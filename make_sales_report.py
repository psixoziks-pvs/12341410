# -*- coding: utf-8 -*-
"""Генерация файла sales_dynamics.xlsx: таблица + линейный график."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.utils import get_column_letter

MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
          "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

# Точные дельты из ТЗ (сумма = 237 363); базовые значения 2020/2021 подобраны так,
# чтобы суммы дали ровно 1 982 603 и 2 219 966, а линии соответствовали описанию графика.
DELTA = [22597, 20616, -2938, 8649, 15243, 29440,
         12766, 43701, 21687, 7628, 34489, 23485]

SALES_2020 = [158830, 164920, 178240, 169540, 161280, 167910,
              175430, 150870, 168750, 172340, 163480, 151013]
SALES_2021 = [v + d for v, d in zip(SALES_2020, DELTA)]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Динамика продаж"
ws.sheet_view.showGridLines = False

# ---------------- стили ----------------
GRAY_HEADER = "B4B4B4"          # RGB(180,180,180)
BORDER_GRAY = "B4B4B4"
thin = Side(style="thin", color=BORDER_GRAY)
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
thick_top = Border(left=thin, right=thin, top=Side(style="double", color=BORDER_GRAY), bottom=thin)

fill_header = PatternFill("solid", fgColor=GRAY_HEADER)

# цвета градиента для колонки "Дельта, руб."
RED      = PatternFill("solid", fgColor="FF6464")   # RGB(255,100,100)
ORANGE   = PatternFill("solid", fgColor="FFA500")   # RGB(255,165,0)
YELLOW   = PatternFill("solid", fgColor="FFFF64")   # RGB(255,255,100)
L_YELLOW = PatternFill("solid", fgColor="FFFACD")   # светло-жёлтый
L_GREEN  = PatternFill("solid", fgColor="90EE90")   # RGB(144,238,144)
GREEN    = PatternFill("solid", fgColor="5FCB5F")   # зелёный
BRIGHT_G = PatternFill("solid", fgColor="00C800")   # RGB(0,200,0)

def delta_fill(d):
    """Правило градиента по величине дельты."""
    if d < 0:
        return RED
    if d < 10000:
        return ORANGE
    if d < 20000:
        return YELLOW
    if d < 30000:
        return L_GREEN
    if d < 40000:
        return GREEN
    return BRIGHT_G

# точные цвета из ТЗ по месяцам (перекрывают общее правило там, где указано)
SPECIFIED_FILLS = {
    0: L_YELLOW,   # Январь  22 597  — светло-жёлтый
    1: L_YELLOW,   # Февраль 20 616  — светло-жёлтый
    2: RED,        # Март    -2 938  — красный
    3: ORANGE,     # Апрель  8 649   — оранжевый
    4: YELLOW,     # Май     15 243  — жёлто-оранжевый -> жёлтый
    5: L_GREEN,    # Июнь    29 440  — светло-зелёный
    6: YELLOW,     # Июль    12 766  — жёлтый
    7: BRIGHT_G,   # Август  43 701  — ярко-зелёный
    8: L_YELLOW,   # Сентябрь 21 687 — светло-жёлтый
    9: ORANGE,     # Октябрь 7 628   — оранжевый
    10: GREEN,     # Ноябрь 34 489   — зелёный
    11: L_YELLOW,  # Декабрь 23 485  — светло-жёлтый
}

NUM_FMT = "# ##0"           # разделитель тысяч пробелом
PCT_FMT = '0%'

font_norm = Font(name="Calibri", size=11, color="000000")
font_bold = Font(name="Calibri", size=11, bold=True, color="000000")
font_title = Font(name="Calibri", size=12, color="000000")

# ---------------- общий заголовок над обоими блоками — по центру (A1:K1) ----------------
ws.merge_cells("A1:K1")
ws["A1"] = "Динамика продаж 2020/2021"
ws["A1"].font = font_title
ws["A1"].alignment = Alignment(horizontal="center")

# ---------------- шапка ----------------
HEADERS = ["Месяц", "2020", "2021", "Δ, руб.", "Δ, %"]
for c, h in enumerate(HEADERS, start=1):
    cell = ws.cell(row=2, column=c, value=h)
    cell.font = font_bold
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = border_all

# ---------------- данные ----------------
total20 = total21 = 0
for i, m in enumerate(MONTHS):
    r = 3 + i
    v20, v21 = SALES_2020[i], SALES_2021[i]
    d = v21 - v20
    total20 += v20
    total21 += v21

    c1 = ws.cell(row=r, column=1, value=m)
    c1.font = font_norm
    c1.alignment = Alignment(horizontal="left")

    c2 = ws.cell(row=r, column=2, value=v20)
    c3 = ws.cell(row=r, column=3, value=v21)
    for cc in (c2, c3):
        cc.font = font_norm
        cc.number_format = NUM_FMT
        cc.alignment = Alignment(horizontal="right")

    c4 = ws.cell(row=r, column=4, value=d)
    c4.font = font_norm
    c4.number_format = NUM_FMT
    c4.alignment = Alignment(horizontal="right")
    c4.fill = SPECIFIED_FILLS.get(i, delta_fill(d))

    c5 = ws.cell(row=r, column=5, value=round(d / v20, 4))
    c5.font = font_norm
    c5.number_format = PCT_FMT
    c5.alignment = Alignment(horizontal="right")

    for c in range(1, 6):
        ws.cell(row=r, column=c).border = border_all

# ---------------- строка "Итого" ----------------
r_total = 15
d_total = total21 - total20
vals = ["Итого", total20, total21, d_total, d_total / total20]
for c, v in enumerate(vals, start=1):
    cell = ws.cell(row=r_total, column=c, value=v)
    cell.font = font_bold
    cell.fill = fill_header
    cell.border = thick_top
    if c == 1:
        cell.alignment = Alignment(horizontal="left")
    else:
        cell.alignment = Alignment(horizontal="right")
        cell.number_format = NUM_FMT
# особые форматы для Итого: Δ руб. с ".0", Δ% без разрядов
ws.cell(row=r_total, column=4).number_format = "# ##0,0"
ws.cell(row=r_total, column=5).number_format = "0%"

# ---------------- размеры колонок ----------------
widths = [12, 11, 11, 11, 8]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ================= ГРАФИК =================
chart = LineChart()
chart.title = "Динамика продаж 2020/2021"
chart.style = 2
chart.height = 13
chart.width = 24
chart.y_axis.title = None
chart.x_axis.title = None
chart.legend.position = "b"
chart.legend.border = None

data = Reference(ws, min_col=2, max_col=3, min_row=2, max_row=14)
cats = Reference(ws, min_col=1, min_row=3, max_row=14)
chart.add_data(data, titles_from_data=True)
# категории — текстовые подписи (ссылка на диапазон с названиями месяцев)
from openpyxl.chart.data_source import AxDataSource, StrRef
for s in chart.series:
    s.cat = AxDataSource(strRef=StrRef(f"'{ws.title}'!$A$3:$A$14"))

BLUE_RGB = "00467F"     # RGB(0,70,127)
ORGB_RGB = "C07832"     # RGB(192,120,50)

for idx, s in enumerate(chart.series):
    s.smooth = False
    color = BLUE_RGB if idx == 0 else ORGB_RGB
    lp = LineProperties(solidFill=color, w=19050)  # ~1.5pt
    s.graphicalProperties = GraphicalProperties(ln=lp)
    mk = Marker(symbol="circle", size=5)
    mk.spPr = GraphicalProperties(solidFill=color)
    s.marker = mk

# Ось X — подписи под углом ~45°, мелкий шрифт
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (RichTextProperties, Paragraph,
                                   ParagraphProperties, CharacterProperties)

chart.x_axis.txPr = RichText(bodyPr=RichTextProperties(rot=-2700000, vert="horz"),
                             p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=800)),
                                          endParaRPr=CharacterProperties(sz=800))])
chart.x_axis.delete = False
chart.y_axis.delete = False
# явная ориентация осей: категории внизу, значения слева (по умолчанию openpyxl ставит catAx axPos="l")
chart.x_axis.axPos = "b"
chart.y_axis.axPos = "l"

# Ось Y: диапазон 120 000 – 200 000, шаг 10 000, формат с пробелом
chart.y_axis.scaling.min = 120000
chart.y_axis.scaling.max = 200000
chart.y_axis.majorUnit = 10000
chart.y_axis.numFmt = "# ##0"

# Сетка: горизонтальные линии светло-серые RGB(220,220,220)
chart.y_axis.majorGridlines.spPr = GraphicalProperties(
    ln=LineProperties(solidFill="DCDCDC", w=9525))
chart.x_axis.majorGridlines = None

# Разместить график справа от таблицы (отступ ~2 колонки, ниже общего заголовка)
ws.add_chart(chart, "G2")

wb.save("/workspace/sales_dynamics.xlsx")
print("OK, totals:", total20, total21, d_total, round(d_total/total20*100, 1), "%")
