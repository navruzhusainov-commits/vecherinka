# -*- coding: utf-8 -*-
"""Генерация Excel-файла «Вложения и планируемый доход» по вечеру Gleb & Noemy в Foster Club.
Данные собраны из презентаций репозитория (Foster_Gleb_Noemy_Presentation.pptx/.pdf и генераторов).
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
F_SUB = Font(name="Calibri", size=11, bold=False, color=WHITE)
F_H = Font(name="Calibri", size=11, bold=True, color=WHITE)
F_B = Font(name="Calibri", size=11, color=DARK)
F_BB = Font(name="Calibri", size=11, bold=True, color=DARK)
F_TOTAL = Font(name="Calibri", size=12, bold=True, color=WHITE)
F_NOTE = Font(name="Calibri", size=10, italic=True, color="6B6558")

FILL_TITLE = PatternFill("solid", fgColor=DARK)
FILL_HEAD = PatternFill("solid", fgColor=RED)
FILL_TOTAL = PatternFill("solid", fgColor=DARK)
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
    c.font = F_TITLE
    c.fill = FILL_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=1, value=subtitle)
    c.font = F_SUB
    c.fill = FILL_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 30


def header_row(ws, row, headers, widths=None):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = F_H
        c.fill = FILL_HEAD
        c.alignment = CENTER
        c.border = BORDER
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 22


def data_row(ws, row, values, bold=False, money_cols=(), fill=None, note_col=None):
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = F_BB if bold else F_B
        c.border = BORDER
        c.alignment = LEFT if i != 2 or 2 not in money_cols else RIGHT
        if i in money_cols:
            c.number_format = MONEY
            c.alignment = RIGHT
        if fill:
            c.fill = fill
    return row + 1


wb = Workbook()

# =====================================================================
# ЛИСТ 1 — СВОДКА
# =====================================================================
ws = wb.active
ws.title = "Сводка"
title_block(ws, 3, "Вложения и планируемый доход — Gleb Filipchenkow & Noemy Abrantes",
            "«Чёрным по красному» • 5 декабря 2026 (сб), 22:00–06:00 • Foster Nightclub, Крым (Mriya Resort & SPA, Ялта, с. Оползневое)")

info = [
    ("Событие", "Первый привоз в Крым — хиты «Mallorca» (Gleb Filipchenkow) и «Garota»/«HOJE» (Noemy Abrantes)"),
    ("Загрузка", "300–350 гостей, цель — SOLD OUT; дресс-код Total Black + красная деталь"),
    ("Билеты", "3 000 ₽ (01–07.11) → 3 500 ₽ (08–14.11) → 4 000 ₽ (с 15.11 / в день события)"),
    ("Старт продаж", "01.11.2026"),
    ("Участники", "FOSTER — площадка (клуб, бар, кухня, персонал, VIP; аренда 0 ₽) • КАКТУСА — организатор (все риски и продажи) • MRIYA — партнёр (проживание артистов без скидки)"),
]
r = 4
for k, v in info:
    ws.cell(row=r, column=1, value=k).font = F_BB
    ws.cell(row=r, column=1).fill = FILL_KEY
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=1).alignment = LEFT
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws.cell(row=r, column=2, value=v)
    c.font = F_B
    c.alignment = WRAP
    c.border = BORDER
    ws.cell(row=r, column=3).border = BORDER
    ws.row_dimensions[r].height = 28
    r += 1

r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
c = ws.cell(row=r, column=1, value="ИТОГОВАЯ СВОДКА ПО СТОРОНАМ")
c.font = F_H
c.fill = FILL_HEAD
c.alignment = CENTER
c.border = BORDER
ws.cell(row=r, column=2).border = BORDER
ws.cell(row=r, column=3).border = BORDER
r += 1
header_row(ws, r, ["Показатель", "Клуб FOSTER (площадка)", "Организатор КАКТУСА"], widths=[26, 46, 52])
r += 1
summary = [
    ("Вложения (затраты)", 284476, "~1 120 040 ₽ — гонорар 550К + логистика 80К + промо 100К + артисты 75К + тех/видео/фото 85К + охрана 30К + райдер 30К + компенсации клубу 200 040 ₽"),
    ("Планируемый доход", 768025, "Билеты: 825 000 ₽ (250 чел.) / 1 020 000 ₽ (300 чел.) / 1 190 000 ₽ (350 чел., средний чек 3 400 ₽)"),
    ("Планируемый результат", 483549, "–295 040 ₽ (250) / –100 040 ₽ (300) / +69 960 ₽ только при SOLD OUT 350"),
    ("Гарант / риск", "330 000 ₽ гарантированы договором (коктейли 180К + столы 50К + VIP 70К) + 20 040 ₽ шаттл; оплата 1–3 дня после события", "Высокий риск — в плюс только при 350 гостях; мотивация продавать до последнего билета"),
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
    c.border = BORDER
    c.alignment = WRAP
    c.font = F_B
    if fill:
        c.fill = fill
    ws.row_dimensions[r].height = 44
    r += 1
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
c = ws.cell(row=r, column=1, value="Источник: презентации репозитория — Foster_Gleb_Noemy_Presentation.pptx/.pdf, FOSTER_Club_презентация.pptx, генераторы create_pptx.py / create_pdf.py / vecherinka/create_final*.py")
c.font = F_NOTE
c.alignment = WRAP

# =====================================================================
# ЛИСТ 2 — ВЛОЖЕНИЯ FOSTER
# =====================================================================
ws = wb.create_sheet("Вложения Foster")
title_block(ws, 3, "Вложения (затраты) клуба FOSTER — всего 284 476 ₽",
            "Только операционка: аренда, артисты, оборудование, проживание — 0 ₽. Все расходы полностью покрываются доходами клуба.")
r = 4
header_row(ws, r, ["Статья затрат", "Сумма", "Примечание"], widths=[34, 16, 56])
r += 1
foster_costs = [
    ("Персонал клуба", 165470, "Официанты, бармены, повара, менеджмент (~58% всех затрат)"),
    ("Себестоимость бара (фудкост)", 92006, "25% от валовой выручки бара 368 025 ₽"),
    ("Себестоимость кухни (фудкост)", 27000, "27% от валовой выручки кухни 100 000 ₽"),
    ("Аренда площадки", 0, "Площадка своя"),
    ("Гонорар артистов", 0, "Платит Кактуса"),
    ("Оборудование / свет / звук", 0, "Уже на месте"),
    ("Проживание артистов (скидка Mriya)", 0, "Отменено по договорённости"),
]
for i, (a, s, n) in enumerate(foster_costs):
    fill = FILL_ALT if i % 2 else None
    r = data_row(ws, r, [a, s, n], money_cols=(2,), fill=fill)
# итог
for i, v in enumerate(["ИТОГО вложения Foster", 284476, "Фактически нулевые инвестиции — только операционка, заложенная в выручку"], start=1):
    c = ws.cell(row=r, column=i, value=v)
    c.font = F_TOTAL
    c.fill = FILL_TOTAL
    c.border = BORDER
    if i == 2:
        c.number_format = MONEY
        c.alignment = RIGHT
    else:
        c.alignment = LEFT
ws.row_dimensions[r].height = 24

# =====================================================================
# ЛИСТ 3 — ВЛОЖЕНИЯ КАКТУСЫ
# =====================================================================
ws = wb.create_sheet("Вложения Кактусы")
title_block(ws, 3, "Вложения организатора КАКТУСА — всего ~1 120 040 ₽",
            "Кактуса берёт все риски: гонорары, логистика, продакшн, продвижение, продажи билетов. 200 040 ₽ из них — фиксированные компенсации клубу Foster.")
r = 4
header_row(ws, r, ["Статья затрат", "Сумма", "Тип / примечание"], widths=[42, 16, 52])
r += 1
kaktus_costs = [
    ("Гонорар хедлайнеров (Gleb Filipchenkow & Noemy Abrantes)", 550000, "Фикс"),
    ("Компенсация клубу за welcome-коктейли (300 × 600 ₽)", 180000, "→ Foster, фикс в договоре, независимо от явки"),
    ("Оплата шаттла (охрана, клининг, хостес Велком-центра)", 20040, "→ Foster, фикс в договоре"),
    ("Перелёт / поезд / трансфер (логистика)", 80000, "Допущение (~)"),
    ("Продвижение, блогеры, афиши (промо)", 100000, "Допущение (~)"),
    ("Местные артисты (перформанс, go-go)", 75000, "Среднее"),
    ("Технический райдер (сверх клуба)", 30000, "Допущение (~)"),
    ("Видеосъёмка и монтаж", 30000, "Продакшн"),
    ("Фотоотчёт", 25000, "Продакшн"),
    ("Внешняя охрана", 30000, "Сверх шаттла (~)"),
]
for i, (a, s, n) in enumerate(kaktus_costs):
    fill = FILL_ALT if i % 2 else None
    r = data_row(ws, r, [a, s, n], money_cols=(2,), fill=fill)
for i, v in enumerate(["ИТОГО вложения Кактусы", 1120040, "Все риски на организаторе; в плюс только при SOLD OUT 350 гостей"], start=1):
    c = ws.cell(row=r, column=i, value=v)
    c.font = F_TOTAL
    c.fill = FILL_TOTAL
    c.border = BORDER
    if i == 2:
        c.number_format = MONEY
        c.alignment = RIGHT
    else:
        c.alignment = LEFT
ws.row_dimensions[r].height = 24
r += 2
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
c = ws.cell(row=r, column=1, value="Из них фиксированные компенсации клубу Foster: 200 040 ₽ = 180 000 ₽ (коктейли) + 20 040 ₽ (шаттл) — не зависят от явки.")
c.font = F_NOTE
c.alignment = WRAP

# =====================================================================
# ЛИСТ 4 — ДОХОДЫ И ПРИБЫЛЬ
# =====================================================================
ws = wb.create_sheet("Доходы и прибыль")
title_block(ws, 5, "Планируемый доход и чистая прибыль — вечер 5 декабря 2026",
            "Доходы клуба Foster 768 025 ₽ + кальяны Макао • сценарии выручки Кактусы от продажи билетов • гарант 330 000 ₽ по договору")

r = 4
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="1) Доходы клуба FOSTER — итого 768 025 ₽ + кальяны Макао (по факту)")
c.font = F_H; c.fill = FILL_HEAD; c.alignment = CENTER
for col in range(2, 6):
    ws.cell(row=r, column=col).border = BORDER
ws.cell(row=r, column=1).border = BORDER
r += 1
header_row(ws, r, ["Источник дохода", "Сумма", "Примечание", "", ""], widths=[38, 16, 44, 14, 16])
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
r += 1
foster_income = [
    ("Бар (валовая выручка)", 368025, ""),
    ("Кухня (валовая выручка)", 100000, ""),
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
for i, v in enumerate(["ИТОГО доходы Foster", 768025, "+ кальяны от Макао — по факту продаж, полностью клубу"], start=1):
    c = ws.cell(row=r, column=i if i < 3 else 3, value=v)
    c.font = F_TOTAL; c.fill = FILL_TOTAL; c.border = BORDER
    if i == 2:
        c.number_format = MONEY; c.alignment = RIGHT
    else:
        c.alignment = LEFT
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
for col in (4, 5):
    ws.cell(row=r, column=col).fill = FILL_TOTAL
    ws.cell(row=r, column=col).border = BORDER
ws.row_dimensions[r].height = 22
r += 2

# --- прибыль Foster ---
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="2) Чистая прибыль клуба FOSTER за вечер")
c.font = F_H; c.fill = FILL_HEAD; c.alignment = CENTER
for col in range(1, 6):
    ws.cell(row=r, column=col).border = BORDER
r += 1
header_row(ws, r, ["Показатель", "Сумма", "Формула / примечание", "", ""])
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
r += 1
profit_rows = [
    ("Доходы", 768025, "Бар 368 025 + кухня 100 000 + коктейли 180 000 + столы 50 000 + VIP 70 000", False),
    ("– Затраты (вложения клуба)", -284476, "Персонал 165 470 + бар 92 006 + кухня 27 000", False),
    ("= Чистая прибыль за вечер", 483549, "768 025 – 284 476 = 483 549 ₽ + кальяны Макао сверху", True),
    ("Рентабельность", 483549 / 768025, "483 549 / 768 025 ≈ 63%", True),
    ("Гарантировано договором (фикс)", 330000, "180 000 коктейли + 50 000 столы + 70 000 VIP №1/№2 — независимо от явки; даже при 200 гостях клуб в плюсе. + 20 040 ₽ шаттл. Оплата 1–3 рабочих дня после события", True),
]
for i, (a, s, n, bold) in enumerate(profit_rows):
    ws.cell(row=r, column=1, value=a).font = F_BB if bold else F_B
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=1).alignment = LEFT
    c = ws.cell(row=r, column=2, value=s)
    c.border = BORDER; c.alignment = RIGHT
    c.font = F_BB if bold else F_B
    c.number_format = PCT if a == "Рентабельность" else MONEY
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    c = ws.cell(row=r, column=3, value=n)
    c.font = F_B; c.alignment = WRAP; c.border = BORDER
    for col in (4, 5):
        ws.cell(row=r, column=col).border = BORDER
    if bold:
        for col in range(1, 6):
            ws.cell(row=r, column=col).fill = FILL_KEY
    ws.row_dimensions[r].height = 30 if len(n) > 60 else 18
    r += 1
r += 1

# --- сценарии Кактусы ---
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="3) Сценарии дохода КАКТУСЫ — продажа билетов (средний чек 3 400 ₽; цены 3 000 → 3 500 → 4 000 ₽)")
c.font = F_H; c.fill = FILL_HEAD; c.alignment = CENTER
for col in range(1, 6):
    ws.cell(row=r, column=col).border = BORDER
r += 1
header_row(ws, r, ["Сценарий", "Гости, чел.", "Средний билет", "Выручка", "Результат (выручка – 1 120 040 ₽)"])
r += 1
scenarios = [
    ("Пессимистичный", 250, 3300, 825000, -295040),
    ("Базовый", 300, 3400, 1020000, -100040),
    ("Оптимистичный / SOLD OUT", 350, 3400, 1190000, 69960),
]
for i, (name, guests, ticket, revenue, result) in enumerate(scenarios):
    fill = FILL_ALT if i % 2 else None
    vals = [name, guests, ticket, revenue, result]
    for col, v in enumerate(vals, start=1):
        c = ws.cell(row=r, column=col, value=v)
        c.border = BORDER
        c.font = F_BB if i == 2 else F_B
        if fill: c.fill = fill
        if col in (3, 4, 5):
            c.number_format = MONEY
            c.alignment = RIGHT
        elif col == 2:
            c.alignment = CENTER
        else:
            c.alignment = LEFT
    if i == 2:
        for col in range(1, 6):
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=RED)
            ws.cell(row=r, column=col).font = Font(name="Calibri", size=11, bold=True, color=WHITE)
    r += 1
r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="Кактуса в плюс только при SOLD OUT 350 гостей → мотивация продавать до последнего билета. Клуб Foster защищён: его 200 040 ₽ компенсаций (180К коктейли + 20К шаттл) зафиксированы в договоре независимо от явки.")
c.font = F_NOTE; c.alignment = WRAP
ws.row_dimensions[r].height = 28
r += 2
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
c = ws.cell(row=r, column=1, value="Следующие шаги: согласовать договор (фикс 330 000 ₽ + компенсации) → старт продаж 01.11.2026 → 5 декабря 2026, SOLD OUT 350 гостей.")
c.font = F_BB; c.alignment = WRAP

# ---------- общие настройки ----------
for sheet in wb.worksheets:
    sheet.sheet_view.showGridLines = False

out = "/home/user/vecherinka/Вложения_и_планируемый_доход.xlsx"
wb.save(out)
print("Сохранено:", out)
