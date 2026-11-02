# -*- coding: utf-8 -*-
"""Генерация Excel-файла «Вложения и планируемый доход» по вечеру Gleb & Noemy в Foster Club.
Данные из презентаций репозитория + уточнения заказчика:
  - аренда клуба 200 000 ₽ и аренда оборудования 250 000 ₽ предоставляются Кактусе БЕСПЛАТНО (экономия 450 000 ₽);
  - персонал — вызов в переработку ~160 000 ₽;
  - бар 368 025 ₽ + кухня 100 000 ₽ = 468 025 ₽ — фактическая выручка прошлого года (бенчмарк).
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------- палитра в духе презентации (красно-чёрный) ----------
RED = "E60613"
DARK = "141419"
LIGHT = "F5F1E8"
GREY = "EDEAE2"
WHITE = "FFFFFF"

thin = Side(style="thin", color="B9B2A4")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

F_TITLE = Font(name="Calibri", size=16, bold=True, color=WHITE)
F_SUB = Font(name="Calibri", size=11, color=WHITE)
F_H = Font(name="Calibri", size=11, bold=True, color=WHITE)
F_B = Font(name="Calibri", size=11, color=DARK)
F_BB = Font(name="Calibri", size=11, bold=True, color=DARK)
F_TOTAL = Font(name="Calibri", size=12, bold=True, color=WHITE)
F_NOTE = Font(name="Calibri", size=10, italic=True, color="6B6558")

FILL_TITLE = PatternFill("solid", fgColor=DARK)
FILL_HEAD = PatternFill("solid", fgColor=RED)
FILL_TOTAL = PatternFill("solid", fgColor=DARK)
FILL_TOTAL_RED = PatternFill("solid", fgColor=RED)
FILL_ALT = PatternFill("solid", fgColor=GREY)
FILL_KEY = PatternFill("solid", fgColor=LIGHT)

MONEY = '#,##0 "₽";[Red]-#,##0 "₽"'
PCT = '0%'

WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def title_block(ws, ncols, title, subtitle):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = F_TITLE; c.fill = FILL_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = F_SUB; c.fill = FILL_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30


def header_row(ws, row, headers, widths=None):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = F_H; c.fill = FILL_HEAD; c.alignment = CENTER; c.border = BORDER
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 22


def data_row(ws, row, values, bold=False, money_cols=(), fill=None):
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = F_BB if bold else F_B
        c.border = BORDER
        c.alignment = RIGHT if i in money_cols else LEFT
        if i in money_cols:
            c.number_format = MONEY
        if fill:
            c.fill = fill
    return row + 1


def total_row(ws, row, values, money_cols=(), red=False):
    fill = FILL_TOTAL_RED if red else FILL_TOTAL
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = F_TOTAL; c.fill = fill; c.border = BORDER
        c.alignment = RIGHT if i in money_cols else LEFT
        if i in money_cols:
            c.number_format = MONEY
    ws.row_dimensions[row].height = 24
    return row + 1


wb = Workbook()

# =====================================================================
# ЛИСТ 1 — СВОДКА
# =====================================================================
ws = wb.active
ws.title = "Сводка"
title_block(ws, 3, "Вложения и планируемый доход — Gleb Filipchenkow & Noemy Abrantes",
            "«Чёрным по красному» • 5 декабря 2026 (сб), 22:00–06:00 • Foster Nightclub, Крым (Mriya Resort & SPA, Ялта, с. Оползневое) • аренда клуба и оборудования для Кактусы — бесплатно")

info = [
    ("Событие", "Первый привоз в Крым — хиты «Mallorca» (Gleb Filipchenkow) и «Garota»/«HOJE» (Noemy Abrantes)"),
    ("Загрузка", "300–350 гостей, цель — SOLD OUT; дресс-код Total Black + красная деталь"),
    ("Билеты", "3 000 ₽ (01–07.11) → 3 500 ₽ (08–14.11) → 4 000 ₽ (с 15.11 / в день события)"),
    ("Старт продаж", "01.11.2026"),
    ("Участники", "FOSTER — площадка (клуб, бар, кухня, персонал, VIP; аренда клуба 200К + оборудование 250К — Кактусе бесплатно) • КАКТУСА — организатор (все риски и продажи) • MRIYA — партнёр (проживание артистов без скидки)"),
    ("Бенчмарк", "Бар 368 025 ₽ + кухня 100 000 ₽ = 468 025 ₽ — фактическая выручка клуба при таком же мероприятии с Кактусой в прошлом году"),
]
r = 4
for k, v in info:
    ws.cell(row=r, column=1, value=k).font = F_BB
    ws.cell(row=r, column=1).fill = FILL_KEY
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=1).alignment = LEFT
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws.cell(row=r, column=2, value=v)
    c.font = F_B; c.alignment = WRAP; c.border = BORDER
    ws.cell(row=r, column=3).border = BORDER
    ws.row_dimensions[r].height = 30
    r += 1

r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
c = ws.cell(row=r, column=1, value="ИТОГОВАЯ СВОДКА ПО СТОРОНАМ")
c.font = F_H; c.fill = FILL_HEAD; c.alignment = CENTER; c.border = BORDER
ws.cell(row=r, column=2).border = BORDER
ws.cell(row=r, column=3).border = BORDER
r += 1
header_row(ws, r, ["Показатель", "Клуб FOSTER (площадка)", "Организатор КАКТУСА"], widths=[26, 50, 52])
r += 1
summary = [
    ("Вложения (затраты)", 729006, "~1 240 040 ₽ — гонорар 550К + логистика 80К + проживание 60К + суточные 40К + бытовой райдер 20К + техрайдер 30К + артисты 75К + промо 100К + видео/фото 55К + охрана 30К + компенсации клубу 200 040 ₽; аренда клуба и оборудования — 0 ₽ (экономия 450 000 ₽)"),
    ("Состав вложений Foster", "279 006 ₽ операционка (персонал ~160К + бар 92К + кухня 27К) + 450 000 ₽ бесплатная аренда клуба (200К) и оборудования (250К); артисты и райдер клубу — 0 ₽", None),
    ("Планируемый доход", 768025, "Билеты: 825 000 ₽ (250 чел.) / 1 020 000 ₽ (300 чел.) / 1 190 000 ₽ (350 × 3 400) / 1 400 000 ₽ (350 × 4 000)"),
    ("Планируемый результат", 489019, "–415 040 ₽ (250) / –220 040 ₽ (300) / –50 040 ₽ (350 × 3 400) / +159 960 ₽ (350 × 4 000)"),
    ("Результат с учётом аренды", 39019, "В плюс только при средней цене билета выше ~3 543 ₽ (весь райдер на Кактусе)"),
    ("Гарант / защита", "330 000 ₽ гарантированы договором (коктейли 180К + столы 50К + VIP 70К) + 20 040 ₽ шаттл + бесплатная аренда/площадка 450К; артисты и райдер — 0 ₽; оплата 1–3 дня после события", "Мотивация поднимать цену к 4 000 ₽ — их 200 040 ₽ компенсаций клубу зафиксированы фиксом, клуб защищён при любой явке"),
]
for i, (k, foster, kaktus) in enumerate(summary):
    fill = FILL_ALT if i % 2 else None
    ws.cell(row=r, column=1, value=k).font = F_BB
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=1).alignment = LEFT
    if fill:
        ws.cell(row=r, column=1).fill = fill
    c = ws.cell(row=r, column=2, value=foster)
    c.border = BORDER
    c.alignment = WRAP if isinstance(foster, str) else RIGHT
    c.font = F_B
    if isinstance(foster, (int, float)):
        c.number_format = MONEY
    if fill:
        c.fill = fill
    c = ws.cell(row=r, column=3, value=kaktus)
    c.border = BORDER; c.alignment = WRAP; c.font = F_B
    if fill:
        c.fill = fill
    ws.row_dimensions[r].height = 46
    r += 1
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
c = ws.cell(row=r, column=1, value="Источник: презентации репозитория + уточнения заказчика (аренда 200К/оборудование 250К бесплатно, персонал ~160К в переработку, выручка прошлого года 468 025 ₽)")
c.font = F_NOTE; c.alignment = WRAP

# =====================================================================
# ЛИСТ 2 — ВЛОЖЕНИЯ FOSTER
# =====================================================================
ws = wb.create_sheet("Вложения Foster")
title_block(ws, 3, "Вложения клуба FOSTER — операционка 279 006 ₽, всего с арендой/площадкой 729 006 ₽",
            "Персонал — вызов в переработку ~160 000 ₽. Аренда клуба (200 000 ₽) и оборудования (250 000 ₽) предоставляются Кактусе БЕСПЛАТНО — это вложение клуба, 0 ₽ для организатора.")
r = 4
header_row(ws, r, ["Статья затрат", "Сумма", "Примечание"], widths=[40, 16, 56])
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
c = ws.cell(row=r, column=1, value="Операционные затраты (денежные)")
c.font = F_H; c.fill = FILL_HEAD; c.alignment = CENTER
for col in range(1, 4):
    ws.cell(row=r, column=col).border = BORDER
r += 1
foster_oper = [
    ("Персонал — вызов в переработку", 160000, "Подготовка и проведение мероприятия, ориентир ~160 тыс ₽ (уточнение вместо 165 470 ₽)"),
    ("Себестоимость бара (фудкост)", 92006, "25% от валовой выручки бара 368 025 ₽"),
    ("Себестоимость кухни (фудкост)", 27000, "27% от валовой выручки кухни 100 000 ₽"),
]
for i, (a, s, n) in enumerate(foster_oper):
    r = data_row(ws, r, [a, s, n], money_cols=(2,), fill=FILL_ALT if i % 2 else None)
r = total_row(ws, r, ["ИТОГО операционные затраты", 279006, "Полностью покрываются доходами клуба"], money_cols=(2,))
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
c = ws.cell(row=r, column=1, value="Предоставлено Кактусе бесплатно (неденежное вложение клуба / упущенная выручка)")
c.font = F_H; c.fill = FILL_HEAD; c.alignment = CENTER
for col in range(1, 4):
    ws.cell(row=r, column=col).border = BORDER
r += 1
foster_free = [
    ("Аренда клуба Foster", 200000, "Рыночная стоимость; для Кактусы — 0 ₽ (бесплатно)"),
    ("Аренда оборудования (свет/звук/сцена)", 250000, "Рыночная стоимость; для Кактусы — 0 ₽ (бесплатно)"),
]
for i, (a, s, n) in enumerate(foster_free):
    r = data_row(ws, r, [a, s, n], money_cols=(2,), fill=FILL_ALT if i % 2 else None)
r = total_row(ws, r, ["ИТОГО бесплатно для Кактусы", 450000, "Экономия организатора = 450 000 ₽"], money_cols=(2,))
r += 1
r = total_row(ws, r, ["ВСЕГО ВЛОЖЕНИЙ FOSTER", 729006, "279 006 ₽ операционка + 450 000 ₽ аренда/площадка и оборудование"], money_cols=(2,), red=True)

# =====================================================================
# ЛИСТ 3 — ВЛОЖЕНИЯ КАКТУСЫ
# =====================================================================
ws = wb.create_sheet("Вложения Кактусы")
title_block(ws, 3, "Вложения организатора КАКТУСА — всего ~1 240 040 ₽ (весь райдер на нём)",
            "Кактуса берёт все риски и ВЕСЬ РАЙДЕР артистов (тех + бытовой, файл райдера ред. 28.10.2026 в репозитории): гонорар 550К с УСН, проезд, проживание Люкс, суточные, гримёрная, CDJ/DJM/ULX-D/подиум. Других артистов привозит сам. Аренда клуба и оборудования — БЕСПЛАТНО от Foster (рыночная 450 000 ₽).")
r = 4
header_row(ws, r, ["Статья затрат", "Сумма", "Источник / примечание"], widths=[52, 16, 52])
r += 1
kaktus_costs = [
    ("Гонорар хедлайнеров (Gleb Filipchenkow & Noemy Abrantes)", 550000, "Фикс с УСН — райдер п.1.1 (05.12.2026, Крым)"),
    ("Компенсация клубу за welcome-коктейли (300 × 600 ₽)", 180000, "→ Foster, фикс в договоре, независимо от явки"),
    ("Оплата шаттла (охрана, клининг, хостес Велком-центра)", 20040, "→ Foster, фикс в договоре"),
    ("Логистика: 2 авиабилета эконом+/бизнес + трансфер (Vito/Zeekr, бизнес-класс)", 80000, "Райдер п.2 — проезд за счёт заказчика (~)"),
    ("Проживание артистов: 2 номера Люкс/Полулюкс 4*, минибар + завтраки, ~2 ночи", 60000, "Райдер п.3 — Mriya без скидки, за счёт Кактусы (~)"),
    ("Суточные артистам: 10 000 ₽/чел/сут × 2 чел × 2 сут", 40000, "Райдер п.2.8 (или питание п.4)"),
    ("Бытовой райдер: гримёрная (вода, сок, сэндвичи, фрукты, вино, пиво, чай/кофе, официант)", 20000, "Райдер п.4–5 — оценка (~)"),
    ("Техрайдер: CDJ-3000 ×2, DJM-A9/V10, ULX-D ×2, in-ear, подиум 3×2 м, мониторы", 30000, "Райдер п.6 — ПОЛНОСТЬЮ на Кактусе, сверх бесплатного оборудования клуба (~)"),
    ("Другие артисты от Кактусы (перформанс, go-go, локальный лайн-ап)", 75000, "За счёт Кактусы, клубу 0 ₽"),
    ("Продвижение, блогеры, афиши (промо)", 100000, "Допущение (~); макет афиши — п.7.3"),
    ("Видеосъёмка и монтаж", 30000, "Продакшн (согласование по п.7.2)"),
    ("Фотоотчёт", 25000, "Продакшн"),
    ("Внешняя охрана (в т.ч. личная охрана артистов, п.5.7)", 30000, "Сверх шаттла (~)"),
    ("Аренда клуба Foster", 0, "Рыночная 200 000 ₽ — БЕСПЛАТНО от Foster"),
    ("Аренда оборудования клуба", 0, "Рыночная 250 000 ₽ — БЕСПЛАТНО от Foster"),
]
for i, (a, s, n) in enumerate(kaktus_costs):
    r = data_row(ws, r, [a, s, n], money_cols=(2,), fill=FILL_ALT if i % 2 else None)
r = total_row(ws, r, ["ИТОГО вложения Кактусы", 1240040, "Весь райдер на организаторе; в плюс только при ср. цене билета выше ~3 543 ₽"], money_cols=(2,), red=True)
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
c = ws.cell(row=r, column=1, value="Экономия за счёт бесплатной аренды клуба и оборудования: 450 000 ₽. Фиксированные компенсации клубу Foster: 200 040 ₽ = 180 000 ₽ (коктейли) + 20 040 ₽ (шаттл). По п.7.12 райдера артистам — 10 билетов танцпол + 2 стола (~34 000 ₽ упущенной выручки от билетов).")
c.font = F_NOTE; c.alignment = WRAP
ws.row_dimensions[r].height = 40

# =====================================================================
# ЛИСТ 4 — ДОХОДЫ И ПРИБЫЛЬ
# =====================================================================
ws = wb.create_sheet("Доходы и прибыль")
title_block(ws, 5, "Планируемый доход и прибыль — вечер 5 декабря 2026",
            "Доходы клуба Foster 768 025 ₽ + кальяны Макао • бар 368 025 + кухня 100 000 = 468 025 ₽ — факт прошлого года • гарант 330 000 ₽ по договору")

r = 4
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="1) Доходы клуба FOSTER — итого 768 025 ₽ + кальяны Макао (по факту)")
c.font = F_H; c.fill = FILL_HEAD; c.alignment = CENTER
for col in range(1, 6):
    ws.cell(row=r, column=col).border = BORDER
r += 1
header_row(ws, r, ["Источник дохода", "Сумма", "Примечание", "", ""], widths=[38, 16, 46, 14, 16])
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
r += 1
foster_income = [
    ("Бар (валовая выручка)", 368025, "ФАКТ прошлого года — такое же мероприятие с Кактусой"),
    ("Кухня (валовая выручка)", 100000, "ФАКТ прошлого года (бар + кухня = 468 025 ₽ — бенчмарк)"),
    ("Welcome-коктейли (компенсация от Кактусы)", 180000, "300 × 600 ₽ — фикс, независимо от явки"),
    ("Депозиты столов", 50000, "4 × 5 000 + 3 × 10 000"),
    ("VIP №1 (депозит)", 40000, ""),
    ("VIP №2 (депозит)", 30000, ""),
]
for i, (a, s, n) in enumerate(foster_income):
    fill = FILL_ALT if i % 2 else None
    ws.cell(row=r, column=1, value=a).font = F_B
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=1).alignment = LEFT
    if fill: ws.cell(row=r, column=1).fill = fill
    c = ws.cell(row=r, column=2, value=s)
    c.number_format = MONEY; c.border = BORDER; c.alignment = RIGHT; c.font = F_B
    if fill: c.fill = fill
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    c = ws.cell(row=r, column=3, value=n)
    c.font = F_B; c.alignment = LEFT; c.border = BORDER
    for col in (4, 5):
        ws.cell(row=r, column=col).border = BORDER
        if fill: ws.cell(row=r, column=col).fill = fill
    if fill: c.fill = fill
    r += 1
# итог доходов
ws.cell(row=r, column=1, value="ИТОГО доходы Foster").font = F_TOTAL
ws.cell(row=r, column=1).fill = FILL_TOTAL; ws.cell(row=r, column=1).border = BORDER
ws.cell(row=r, column=1).alignment = LEFT
c = ws.cell(row=r, column=2, value=768025)
c.number_format = MONEY; c.font = F_TOTAL; c.fill = FILL_TOTAL; c.border = BORDER; c.alignment = RIGHT
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
c = ws.cell(row=r, column=3, value="+ кальяны от Макао — по факту продаж, полностью клубу")
c.font = F_TOTAL; c.fill = FILL_TOTAL; c.border = BORDER; c.alignment = LEFT
for col in (4, 5):
    ws.cell(row=r, column=col).fill = FILL_TOTAL; ws.cell(row=r, column=col).border = BORDER
ws.row_dimensions[r].height = 22
r += 2

# --- прибыль Foster ---
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="2) Прибыль клуба FOSTER за вечер")
c.font = F_H; c.fill = FILL_HEAD; c.alignment = CENTER
for col in range(1, 6):
    ws.cell(row=r, column=col).border = BORDER
r += 1
header_row(ws, r, ["Показатель", "Сумма", "Формула / примечание", "", ""])
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
r += 1
profit_rows = [
    ("Доходы", 768025, "Бар 368 025 + кухня 100 000 (факт прошлого года) + коктейли 180 000 + столы 50 000 + VIP 70 000", False, MONEY),
    ("– Операционные затраты", -279006, "Персонал ~160 000 (переработка) + бар 92 006 + кухня 27 000", False, MONEY),
    ("= Операционная прибыль", 489019, "768 025 – 279 006 = 489 019 ₽ + кальяны Макао сверху", True, MONEY),
    ("Рентабельность (операционная)", 489019 / 768025, "489 019 / 768 025 ≈ 64%", True, PCT),
    ("– Бесплатная аренда клуба + оборудования", -450000, "Рыночная 200 000 + 250 000 ₽ — Кактусе бесплатно (вложение клуба)", False, MONEY),
    ("= Чистая прибыль с учётом аренды/площадки", 39019, "768 025 – 729 006 = 39 019 ₽ — клуб всё равно в плюсе", True, MONEY),
    ("Гарантировано договором (фикс)", 330000, "180 000 коктейли + 50 000 столы + 70 000 VIP №1/№2 — независимо от явки; даже при 200 гостях в плюсе (330 000 – 279 006 = +50 994 ₽). + 20 040 ₽ шаттл. Оплата 1–3 рабочих дня после события", True, MONEY),
]
for i, (a, s, n, bold, fmt) in enumerate(profit_rows):
    ws.cell(row=r, column=1, value=a).font = F_BB if bold else F_B
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=1).alignment = LEFT
    c = ws.cell(row=r, column=2, value=s)
    c.border = BORDER; c.alignment = RIGHT
    c.font = F_BB if bold else F_B
    c.number_format = fmt
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    c = ws.cell(row=r, column=3, value=n)
    c.font = F_B; c.alignment = WRAP; c.border = BORDER
    for col in (4, 5):
        ws.cell(row=r, column=col).border = BORDER
    if bold:
        for col in range(1, 6):
            ws.cell(row=r, column=col).fill = FILL_KEY
    ws.row_dimensions[r].height = 34 if len(n) > 70 else 18
    r += 1
r += 1

# --- сценарии Кактусы ---
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="3) Сценарии дохода КАКТУСЫ — продажа билетов (затраты 1 240 040 ₽; цены 3 000 → 3 500 → 4 000 ₽; аренда и оборудование — бесплатно)")
c.font = F_H; c.fill = FILL_HEAD; c.alignment = CENTER
for col in range(1, 6):
    ws.cell(row=r, column=col).border = BORDER
r += 1
header_row(ws, r, ["Сценарий", "Гости, чел.", "Средний билет", "Выручка", "Результат (выручка – 1 240 040 ₽)"])
r += 1
scenarios = [
    ("Пессимистичный", 250, 3300, 825000, -415040),
    ("Базовый", 300, 3400, 1020000, -220040),
    ("SOLD OUT × средняя цена 3 400 ₽", 350, 3400, 1190000, -50040),
    ("SOLD OUT × финальная цена 4 000 ₽", 350, 4000, 1400000, 159960),
]
for i, (name, guests, ticket, revenue, result) in enumerate(scenarios):
    fill = FILL_ALT if i % 2 else None
    vals = [name, guests, ticket, revenue, result]
    for col, v in enumerate(vals, start=1):
        c = ws.cell(row=r, column=col, value=v)
        c.border = BORDER
        c.font = F_BB if i == 3 else F_B
        if fill: c.fill = fill
        if col in (3, 4, 5):
            c.number_format = MONEY
            c.alignment = RIGHT
        elif col == 2:
            c.alignment = CENTER
        else:
            c.alignment = LEFT
    if i == 3:
        for col in range(1, 6):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=RED)
            ws.cell(row=r, column=col).font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    r += 1
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="Весь райдер на Кактусе: при средней цене 3 400 ₽ он в минусе даже при SOLD OUT 350 (–50 040 ₽); в плюс выходит при средней цене выше ~3 543 ₽/билет, т.е. при продаже по финальной цене 4 000 ₽ (+159 960 ₽). Клуб Foster защищён: 200 040 ₽ компенсаций зафиксированы в договоре независимо от явки; аренда и оборудование — бесплатно (экономия 450 000 ₽).")
c.font = F_NOTE; c.alignment = WRAP
ws.row_dimensions[r].height = 28
r += 2
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="Следующие шаги: согласовать договор (фикс 330 000 ₽ + компенсации + бесплатная аренда/площадка 450К) → старт продаж 01.11.2026 → 5 декабря 2026, SOLD OUT 350 гостей.")
c.font = F_BB; c.alignment = WRAP

# ---------- общие настройки ----------
for sheet in wb.worksheets:
    sheet.sheet_view.showGridLines = False

out = "/home/user/vecherinka/Вложения_и_планируемый_доход.xlsx"
wb.save(out)
print("Сохранено:", out)
