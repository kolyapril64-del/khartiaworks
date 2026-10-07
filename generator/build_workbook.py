"""Генератор робочої книги «Добова доповідь ППО» (Excel, .xlsx).

Кожен підрозділ веде свій аркуш із журналом подій (один рядок = одна подія),
а аркуші «Зведення», «За типами цілей» і «Розрахунок» рахуються автоматично.

Використання:
    python build_workbook.py config.json output.xlsx

Формат config.json див. у config.example.json.
"""

import json
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------- розміри
UNIT_SLOTS = 15      # скільки підрозділів вміщує зведення
TYPE_SLOTS = 30      # типів цілей у довіднику
MEANS_SLOTS = 15     # засобів ураження
RESULT_SLOTS = 10
STATUS_SLOTS = 8
STRIKE_SLOTS = 5
ENEMY_SLOTS = 9      # засобів противника (стовпці у зведенні ударів)
RADAR_SLOTS = 8

# Аркуш підрозділу: фіксовані адреси, на які посилається зведення
J_HEAD, J_EXAMPLE, J_FIRST, J_LAST = 34, 35, 36, 335   # журнал цілей
S_LAST = 135                                           # удари противника
RADAR_ROWS = range(11, 15)
RADAR_TOTAL = 15
CREW_ROWS = range(19, 31)
CREW_TOTAL = 31

# Службові значення (на них спираються формули — не перейменовувати)
DESTROYED = "Знищено"
PATROL = "Без цілі (патрулювання)"
G_FPV, G_SMALL, G_OTHER = "ФПВ", "Стр.зб", "Інші"
READY = "боєготові"
STRIKE_CODES = [("РУ", "ракетний удар"), ("АУ", "авіаційний удар"), ("УДК", "удар дроном-камікадзе")]

DEFAULT_TARGETS = [
    "Shahed", "Гербера", "Ланцет", "Молнія", "FPV", "FPV (оптоволокно)", "Бомбер",
    "Mavic / Autel", "DJI Matrice", "Орлан-10/30", "Supercam", "Скат", "Zala", "V2U",
    "НВТ", "Італмас", "КВО (Князь Віщий Олег)", "Сокіл", "Кощей", "Merlin", "Куб",
    "Легіонер", "Аеростат", PATROL,
]
DEFAULT_MEANS = [
    ("FPV-перехоплювач", G_FPV), ("Стрілецька зброя", G_SMALL), ("Скид котушки", G_OTHER),
    ("Переріз оптоволокна", G_OTHER), ("ЗКР 9М37", G_OTHER), ("ПЗРК", G_OTHER),
    ("РЕБ", G_OTHER), ("Інше", G_OTHER),
]
DEFAULT_RESULTS = [DESTROYED, "Пошкоджено", "Не знищено", "Ціль не виявлено, борт повернули",
                   "Борт втрачено", "Патрулювання"]
DEFAULT_STATUS = [READY, "обмежено боєготові", "відновлення", "ремонт", "ротація"]
DEFAULT_ENEMY = ["Шахед / Гербера", "Ланцет", "Молнія", "Італмас", "Куб", "FPV", "КАБ", "Ракета", "Інше"]

# ---------------------------------------------------------------- стиль
FONT = "Arial"
C_NAVY, C_TEAL, C_SLATE = "1F3A5F", "0F6E73", "44546A"
C_INPUT, C_AUTO, C_TILE, C_EXAMPLE = "FFF7D6", "EEF0F3", "F3F6FA", "F7F7F7"
C_LINE = "C9CED6"

THIN = Side(style="thin", color=C_LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM = '0;-0;"–"'
PCT = '0%;-0%;"–"'


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def font(size=10, bold=False, color="1A1A1A", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def q(sheet):
    """Назва аркуша в лапках для посилань."""
    return "'" + sheet.replace("'", "''") + "'"


def cells(ws, rng):
    for row in ws[rng] if ":" in rng else [[ws[rng]]]:
        for c in row:
            yield c


def style(ws, rng, *, f=None, fnt=None, border=BOX, align=None, fmt=None, unlocked=False, merge=False):
    for c in cells(ws, rng):
        if f:
            c.fill = f
        if fnt:
            c.font = fnt
        if border:
            c.border = border
        if align:
            c.alignment = align
        if fmt:
            c.number_format = fmt
        if unlocked:
            c.protection = Protection(locked=False)
    if merge and ":" in rng:
        ws.merge_cells(rng)


CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)


def title(ws, rng, text, size=14):
    ws[rng.split(":")[0]] = text
    style(ws, rng, f=fill(C_NAVY), fnt=font(size, True, "FFFFFF"), border=None,
          align=Alignment(horizontal="left", vertical="center", indent=1), merge=True)


def section(ws, rng, text, color=C_TEAL):
    ws[rng.split(":")[0]] = text
    style(ws, rng, f=fill(color), fnt=font(10, True, "FFFFFF"), border=None,
          align=Alignment(horizontal="left", vertical="center", indent=1), merge=":" in rng)


def header(ws, rng, text=None):
    if text is not None:
        ws[rng.split(":")[0]] = text
    style(ws, rng, f=fill(C_SLATE), fnt=font(9, True, "FFFFFF"), align=CENTER, merge=":" in rng)


def label(ws, rng, text):
    ws[rng.split(":")[0]] = text
    style(ws, rng, f=fill(C_AUTO), fnt=font(9, True, C_SLATE), align=LEFT, merge=":" in rng)


def inp(ws, rng, value=None, fmt=None, align=LEFT):
    if value is not None:
        ws[rng.split(":")[0]] = value
    style(ws, rng, f=fill(C_INPUT), fnt=font(), align=align, fmt=fmt, unlocked=True, merge=":" in rng)


def auto(ws, rng, value=None, fmt=None, bold=False, align=CENTER):
    if value is not None:
        ws[rng.split(":")[0]] = value
    style(ws, rng, f=fill(C_AUTO), fnt=font(10, bold), align=align, fmt=fmt, merge=":" in rng)


def tile(ws, lab_rng, val_rng, text, formula, fmt=NUM):
    ws[lab_rng.split(":")[0]] = text
    style(ws, lab_rng, f=fill(C_TILE), fnt=font(8, True, "5B6B7F"), align=CENTER,
          border=Border(left=THIN, right=THIN, top=THIN), merge=":" in lab_rng)
    ws[val_rng.split(":")[0]] = formula
    style(ws, val_rng, f=fill(C_TILE), fnt=font(18, True, C_NAVY), align=CENTER, fmt=fmt,
          border=Border(left=THIN, right=THIN, bottom=THIN), merge=":" in val_rng)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def list_dv(ws, name, ranges, prompt=None):
    dv = DataValidation(type="list", formula1=name, allow_blank=True, showDropDown=False)
    dv.error, dv.errorTitle = "Оберіть значення зі списку (Довідники).", "Невірне значення"
    if prompt:
        dv.prompt, dv.showInputMessage = prompt, True
    ws.add_data_validation(dv)
    for r in ranges:
        dv.add(r)


def num_dv(ws, ranges, hi=999):
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2=str(hi), allow_blank=True)
    dv.error, dv.errorTitle = f"Ціле число від 0 до {hi}.", "Невірне значення"
    ws.add_data_validation(dv)
    for r in ranges:
        dv.add(r)


def time_dv(ws, ranges):
    dv = DataValidation(type="time", operator="between", formula1="0", formula2="0.999305555555556",
                        allow_blank=True, errorStyle="warning")
    dv.error, dv.errorTitle = "Вводьте час у форматі ГГ:ХХ, напр. 07:24.", "Формат часу"
    ws.add_data_validation(dv)
    for r in ranges:
        dv.add(r)


def date_dv(ws, ranges):
    dv = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                        errorStyle="warning")
    dv.error, dv.errorTitle = "Вводьте дату у форматі ДД.ММ.РРРР.", "Формат дати"
    ws.add_data_validation(dv)
    for r in ranges:
        dv.add(r)


def protect(ws):
    p = ws.protection
    p.sheet = True
    p.autoFilter = False      # фільтри дозволені
    p.formatColumns = False   # ширину стовпців можна змінювати
    p.formatRows = False


def page(ws, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False


def hhmm(ref):
    return f'RIGHT("0"&HOUR({ref}),2)&":"&RIGHT("0"&MINUTE({ref}),2)'


# ================================================================ Довідники
REF = "Довідники"


def ref_range(col, n):
    return f"{q(REF)}!${col}$5:${col}${4 + n}"


def build_reference(wb, cfg):
    ws = wb.create_sheet(REF)
    ws.sheet_properties.tabColor = "8C96A3"
    page(ws)
    title(ws, "A1:T1", "ДОВІДНИКИ — списки для випадаючих меню")
    label(ws, "A2", "Назва об'єднання")
    inp(ws, "B2:F2", cfg["org_name"])
    ws["H2"] = ("Можна дописувати нові значення в порожні жовті клітинки. "
                "Виділені жирним — службові: на них спираються формули, не змінюйте їх.")
    style(ws, "H2:T2", fnt=font(9, italic=True, color="5B6B7F"), border=None, align=LEFT, merge=True)
    ws.row_dimensions[2].height = 30

    def column(col, head, values, slots, service=(), note_col=None, notes=None, note_head=None):
        header(ws, f"{col}4", head)
        if note_col:
            header(ws, f"{note_col}4", note_head)
        for i in range(slots):
            c = f"{col}{5 + i}"
            v = values[i] if i < len(values) else None
            if v in service:
                auto(ws, c, v, bold=True, align=LEFT)
                ws[c].comment = Comment("Службове значення — використовується у формулах.", "ППО")
            else:
                inp(ws, c, v)
            if note_col:
                nv = notes[i] if notes and i < len(notes) else None
                (auto if v in service else inp)(ws, f"{note_col}{5 + i}", nv, align=LEFT)

    units = cfg["units"]
    column("A", "Код аркуша", [u["code"] for u in units], UNIT_SLOTS,
           note_col="B", notes=[u["name"] for u in units], note_head="Повна назва підрозділу")
    column("D", "Типи цілей", DEFAULT_TARGETS, TYPE_SLOTS, service=(PATROL,))
    column("F", "Засоби ураження", [m for m, _ in DEFAULT_MEANS], MEANS_SLOTS,
           note_col="G", notes=[g for _, g in DEFAULT_MEANS], note_head="Група")
    column("I", "Результати", DEFAULT_RESULTS, RESULT_SLOTS, service=(DESTROYED,))
    column("K", "Стан екіпажу", DEFAULT_STATUS, STATUS_SLOTS, service=(READY,))
    column("M", "Тип удару", [c for c, _ in STRIKE_CODES], STRIKE_SLOTS,
           service=tuple(c for c, _ in STRIKE_CODES),
           note_col="N", notes=[d for _, d in STRIKE_CODES], note_head="Розшифровка")
    column("P", "Засоби противника", DEFAULT_ENEMY, ENEMY_SLOTS)
    column("R", "Типи РЛС", cfg.get("radars", []), RADAR_SLOTS)
    column("T", "Групи засобів", [G_FPV, G_SMALL, G_OTHER], 3, service=(G_FPV, G_SMALL, G_OTHER))

    widths(ws, {"A": 12, "B": 30, "C": 2, "D": 26, "E": 2, "F": 24, "G": 10, "H": 2, "I": 30,
                "J": 2, "K": 20, "L": 2, "M": 10, "N": 24, "O": 2, "P": 20, "Q": 2, "R": 20,
                "S": 2, "T": 14})

    for name, col, n in [("lst_units", "A", UNIT_SLOTS), ("lst_targets", "D", TYPE_SLOTS),
                         ("lst_means", "F", MEANS_SLOTS), ("lst_groups", "T", 3),
                         ("lst_results", "I", RESULT_SLOTS), ("lst_status", "K", STATUS_SLOTS),
                         ("lst_strike", "M", STRIKE_SLOTS), ("lst_enemy", "P", ENEMY_SLOTS),
                         ("lst_radar", "R", RADAR_SLOTS)]:
        wb.defined_names[name] = DefinedName(name, attr_text=ref_range(col, n))
    list_dv(ws, "lst_groups", [f"G5:G{4 + MEANS_SLOTS}"])
    ws.freeze_panes = "A5"
    protect(ws)
    return ws


# ================================================================ аркуш підрозділу
def build_unit(wb, unit, idx, radars):
    ws = wb.create_sheet(unit["code"])
    ws.sheet_properties.tabColor = C_TEAL
    page(ws)
    widths(ws, {"A": 6, "B": 8, "C": 18, "D": 15, "E": 15, "F": 15, "G": 16, "H": 16, "I": 22,
                "J": 11, "K": 10, "L": 26, "M": 70, "N": 8, "O": 2, "P": 6, "Q": 8, "R": 9,
                "S": 16, "T": 7, "U": 22, "V": 9, "W": 44})

    # --- шапка
    title(ws, "A1:M1", '="ДОБОВА ДОПОВІДЬ ППО — "&C2')
    ws.row_dimensions[1].height = 26
    label(ws, "A2:B2", "Підрозділ")
    inp(ws, "C2:F2", unit["name"])
    label(ws, "H2", "Дата звіту")
    auto(ws, "I2", f"=IF({q(SUM_SHEET)}!$C$3=\"\",\"\",{q(SUM_SHEET)}!$C$3)", fmt="dd.mm.yyyy", bold=True)
    ws["J2"] = "← береться з аркуша «Зведення»"
    style(ws, "J2:M2", fnt=font(8, italic=True, color="5B6B7F"), border=None, align=LEFT, merge=True)
    label(ws, "A3:B3", "Зміна з")
    inp(ws, "C3", fmt="hh:mm", align=CENTER)
    label(ws, "D3", "по")
    inp(ws, "E3", fmt="hh:mm", align=CENTER)
    label(ws, "H3", "Черговий")
    inp(ws, "I3:M3")
    time_dv(ws, ["C3", "E3"])

    # --- підсумок (авто)
    jr = lambda col: f"${col}${J_FIRST}:${col}${J_LAST}"
    sr = lambda col: f"${col}${J_FIRST}:${col}${S_LAST}"
    section(ws, "A5:M5", "ПІДСУМОК ЗА ДОБУ — рахується автоматично з журналу")
    tile(ws, "A6:B6", "A7:B7", "Виявлено цілей",
         f'=COUNTIFS({jr("C")},"?*",{jr("C")},"<>{PATROL}")')
    tile(ws, "C6:D6", "C7:D7", "Застосувань", f'=COUNTIFS({jr("F")},"?*")')
    tile(ws, "E6:F6", "E7:F7", "Знищено", f'=COUNTIFS({jr("I")},"{DESTROYED}")')
    tile(ws, "G6:H6", "G7:H7", "Ефективність", '=IF(C7=0,"–",E7/C7)', fmt=PCT)
    tile(ws, "I6:J6", "I7:J7", "Витрачено бортів", f'=SUM({jr("J")})')
    tile(ws, "K6:L6", "K7:L7", "Втрачено бортів", f'=SUM({jr("K")})')
    section(ws, "P5:W5", "УДАРИ ПРОТИВНИКА — підсумок", color=C_SLATE)
    tile(ws, "P6:Q6", "P7:Q7", "РУ", f'=COUNTIFS({sr("R")},"РУ")')
    tile(ws, "R6:S6", "R7:S7", "АУ", f'=COUNTIFS({sr("R")},"АУ")')
    tile(ws, "T6:U6", "T7:U7", "УДК", f'=COUNTIFS({sr("R")},"УДК")')
    tile(ws, "V6", "V7", "Втрати о/с", f'=SUM({sr("V")})')
    tile(ws, "W6", "W7", "Засобів противника, од.", f'=SUM({sr("T")})')
    ws.row_dimensions[7].height = 32

    # --- РЛС: наявність і робота
    section(ws, "A9:F9", "СИЛИ І ЗАСОБИ — РЛС")
    section(ws, "H9:M9", "РОБОТА РЛС ЗА ДОБУ")
    header(ws, "A10:B10", "Тип РЛС")
    for col, t in zip("CDEF", ["Всього", "Справні", "На ремонті", "На складі"]):
        header(ws, f"{col}10", t)
    for col, t in zip("HIJKLM", ["Позивний РЛС", "Тип РЛС", "Виявлено цілей", "Включень",
                                 "Час роботи (год:хв)", "Періоди готовності №1 (з – по)"]):
        header(ws, f"{col}10", t)
    ws.row_dimensions[10].height = 28
    for i, r in enumerate(RADAR_ROWS):
        inp(ws, f"A{r}:B{r}", radars[i] if i < len(radars) else None)
        for col in "CDEF":
            inp(ws, f"{col}{r}", fmt="0", align=CENTER)
        inp(ws, f"H{r}")
        inp(ws, f"I{r}")
        inp(ws, f"J{r}", fmt="0", align=CENTER)
        inp(ws, f"K{r}", fmt="0", align=CENTER)
        inp(ws, f"L{r}", fmt="[h]:mm", align=CENTER)
        inp(ws, f"M{r}")
    t = RADAR_TOTAL
    label(ws, f"A{t}:B{t}", "Разом")
    for col in "CDEF":
        auto(ws, f"{col}{t}", f"=SUM({col}{RADAR_ROWS[0]}:{col}{RADAR_ROWS[-1]})", fmt=NUM, bold=True)
    label(ws, f"H{t}:I{t}", "Разом")
    auto(ws, f"J{t}", f"=SUM(J{RADAR_ROWS[0]}:J{RADAR_ROWS[-1]})", fmt=NUM, bold=True)
    auto(ws, f"K{t}", f"=SUM(K{RADAR_ROWS[0]}:K{RADAR_ROWS[-1]})", fmt=NUM, bold=True)
    auto(ws, f"L{t}", f"=SUM(L{RADAR_ROWS[0]}:L{RADAR_ROWS[-1]})", fmt="[h]:mm", bold=True)
    auto(ws, f"M{t}")
    rr = f"{RADAR_ROWS[0]}:{RADAR_ROWS[-1]}".split(":")
    list_dv(ws, "lst_radar", [f"A{rr[0]}:A{rr[1]}", f"I{rr[0]}:I{rr[1]}"])
    num_dv(ws, [f"C{rr[0]}:F{rr[1]}", f"J{rr[0]}:K{rr[1]}"])

    # --- екіпажі
    section(ws, "A17:L17", "ЕКІПАЖІ ПЕРЕХОПЛЮВАЧІВ — НАЯВНІСТЬ")
    header(ws, "A18:B18", "Екіпаж (позивний)")
    for col, t in zip("CDEFGHIJ", ["Тип перехоплювача", "Бортів на позиції (день)",
                                   "Бортів на позиції (ніч)", "Бортів на складі", "Швидкість, км/год",
                                   "Стан", "Район (Н.П.)", "Дата відновлення"]):
        header(ws, f"{col}18", t)
    header(ws, "K18:L18", "Примітка")
    ws.row_dimensions[18].height = 28
    for r in CREW_ROWS:
        inp(ws, f"A{r}:B{r}")
        inp(ws, f"C{r}")
        for col in "DEF":
            inp(ws, f"{col}{r}", fmt="0", align=CENTER)
        inp(ws, f"G{r}", align=CENTER)
        inp(ws, f"H{r}")
        inp(ws, f"I{r}")
        inp(ws, f"J{r}", fmt="dd.mm.yyyy", align=CENTER)
        inp(ws, f"K{r}:L{r}")
    c0, c1, t = CREW_ROWS[0], CREW_ROWS[-1], CREW_TOTAL
    label(ws, f"A{t}:B{t}", "Разом")
    auto(ws, f"C{t}", f'=COUNTA(A{c0}:A{c1})', fmt='0" екіпаж."', bold=True)
    for col in "DEF":
        auto(ws, f"{col}{t}", f"=SUM({col}{c0}:{col}{c1})", fmt=NUM, bold=True)
    auto(ws, f"G{t}")
    auto(ws, f"H{t}", f'=COUNTIFS(H{c0}:H{c1},"{READY}")', fmt='0" боєгот."', bold=True)
    for col in "IJ":
        auto(ws, f"{col}{t}")
    auto(ws, f"K{t}:L{t}")
    list_dv(ws, "lst_status", [f"H{c0}:H{c1}"])
    num_dv(ws, [f"D{c0}:F{c1}"])
    date_dv(ws, [f"J{c0}:J{c1}"])

    # --- журнал цілей
    section(ws, f"A{J_HEAD - 1}:N{J_HEAD - 1}",
            "ЖУРНАЛ ЦІЛЕЙ — один рядок = одна подія (виявлення, застосування, патрулювання)")
    section(ws, f"P{J_HEAD - 1}:W{J_HEAD - 1}", "УДАРИ ПРОТИВНИКА (РУ / АУ / УДК)", color=C_SLATE)
    jcols = ["№", "Час", "Тип цілі", "Н.П. / район", "Квадрат (MGRS)", "Засіб ураження",
             "Модель борта / зброя", "Екіпаж / позиція", "Результат", "Витрачено бортів",
             "Втрачено бортів", "Інші витрати / примітка", "Доповідь (формується автоматично)",
             "Група засобу"]
    scols = ["№", "Час", "Тип удару", "Засіб противника", "К-сть, од.", "Об'єкт / район",
             "Втрати о/с", "Опис / наслідки"]
    for i, t in enumerate(jcols):
        header(ws, f"{get_column_letter(1 + i)}{J_HEAD}", t)
    for i, t in enumerate(scols):
        header(ws, f"{get_column_letter(16 + i)}{J_HEAD}", t)
    ws.row_dimensions[J_HEAD].height = 30

    def report(r):
        p = [
            f'IF(B{r}="","",IFERROR({hhmm(f"B{r}")},B{r})&" ")',
            f'IF(H{r}="","","екіпаж «"&H{r}&"» ")',
            "$C$2",
            f'IF(D{r}="","",", в р-ні н.п. "&D{r})',
            f'IF(E{r}="",""," ("&E{r}&")")',
            f'IF(C{r}="{PATROL}",". Проведено патрулювання",". Виявлено БпЛА «"&C{r}&"»")',
            f'IF(F{r}="","",". Застосовано: "&F{r}&IF(G{r}="",""," «"&G{r}&"»"))',
            f'IF(I{r}="","",". "&I{r})',
            f'IF(N(J{r})>0,". Витрата: "&J{r}&" борт(и)","")',
            f'IF(N(K{r})>0,". Втрати: "&K{r}&" борт(и)","")',
            f'IF(L{r}="",".",". "&L{r}&IF(RIGHT(L{r},1)=".","","."))',
        ]
        return f'=IF(C{r}="","",' + "&".join(p) + ")"

    group = (lambda r: f'=IF(F{r}="","",IFERROR(INDEX({ref_range("G", MEANS_SLOTS)},'
                       f'MATCH(F{r},{ref_range("F", MEANS_SLOTS)},0)),"{G_OTHER}"))')

    for r in range(J_EXAMPLE, J_LAST + 1):
        ex = r == J_EXAMPLE
        auto(ws, f"A{r}", "Приклад" if ex else f'=IF(C{r}="","",COUNTA($C${J_FIRST}:C{r}))')
        inp(ws, f"B{r}", fmt="hh:mm", align=CENTER)
        for col in "CDEFGHI":
            inp(ws, f"{col}{r}")
        inp(ws, f"J{r}", fmt="0", align=CENTER)
        inp(ws, f"K{r}", fmt="0", align=CENTER)
        inp(ws, f"L{r}")
        auto(ws, f"M{r}", report(r), align=LEFT)
        auto(ws, f"N{r}", group(r))
        if r <= S_LAST:
            auto(ws, f"P{r}", "Приклад" if ex else f'=IF(R{r}="","",COUNTA($R${J_FIRST}:R{r}))')
            inp(ws, f"Q{r}", fmt="hh:mm", align=CENTER)
            inp(ws, f"R{r}", align=CENTER)
            inp(ws, f"S{r}")
            inp(ws, f"T{r}", fmt="0", align=CENTER)
            inp(ws, f"U{r}")
            inp(ws, f"V{r}", fmt="0", align=CENTER)
            inp(ws, f"W{r}")

    # рядок-приклад (не враховується у підсумках)
    e = J_EXAMPLE
    for c, v in {"B": 0.30833333, "C": "Zala", "D": "Н.П. (приклад)", "E": "37U XX 00000 00000",
                 "F": "FPV-перехоплювач", "G": "модель борта", "H": "Позивний", "I": DESTROYED,
                 "J": 1, "K": 0, "L": "ЕД – 1 од.", "Q": 0.51319444, "R": "УДК", "S": "Молнія",
                 "T": 1, "U": "позиція (приклад)", "V": 0, "W": "Без втрат."}.items():
        ws[f"{c}{e}"] = v
    for c in [*"ABCDEFGHIJKLMN", *"PQRSTUVW"]:
        ws[f"{c}{e}"].font = font(9, italic=True, color="7F7F7F")
        ws[f"{c}{e}"].fill = fill(C_EXAMPLE)
    ws[f"A{e}"].comment = Comment("Рядок-приклад: не враховується у підсумках. Можна очистити.", "ППО")

    ws.add_table(Table(displayName=f"tblJournal{idx}", ref=f"A{J_HEAD}:N{J_LAST}",
                       tableStyleInfo=TableStyleInfo(name="TableStyleLight1", showRowStripes=False)))
    ws.add_table(Table(displayName=f"tblStrikes{idx}", ref=f"P{J_HEAD}:W{S_LAST}",
                       tableStyleInfo=TableStyleInfo(name="TableStyleLight1", showRowStripes=False)))

    rng = lambda c, last=J_LAST: f"{c}{J_EXAMPLE}:{c}{last}"
    list_dv(ws, "lst_targets", [rng("C")])
    list_dv(ws, "lst_means", [rng("F")])
    list_dv(ws, "lst_results", [rng("I")])
    num_dv(ws, [rng("J"), rng("K")], hi=99)
    time_dv(ws, [rng("B"), rng("Q", S_LAST)])
    list_dv(ws, "lst_strike", [rng("R", S_LAST)], prompt="РУ — ракетний, АУ — авіаційний, УДК — дрон-камікадзе")
    list_dv(ws, "lst_enemy", [rng("S", S_LAST)])
    num_dv(ws, [rng("T", S_LAST), rng("V", S_LAST)])

    res = f"I{J_FIRST}:I{J_LAST}"
    for val, bg, fg in [(DESTROYED, "D8F0DE", "1E6B34"), ("Пошкоджено", "FFF1C2", "7A5A00"),
                        ("Не знищено", "FAD9DC", "8E1B26"), ("Борт втрачено", "FAD9DC", "8E1B26")]:
        ws.conditional_formatting.add(res, CellIsRule(operator="equal", formula=[f'"{val}"'],
                                                      fill=fill(bg), font=Font(name=FONT, color=fg, bold=True)))
    protect(ws)
    return ws


# ================================================================ Зведення
SUM_SHEET = "Зведення"
U_FIRST = 11                     # перший рядок підрозділу у таблиці «Бойова робота»
U_LAST = U_FIRST + UNIT_SLOTS - 1


def ind(code, rng):
    return f'INDIRECT("\'"&{code}&"\'!{rng}")'


def build_summary(wb):
    ws = wb.create_sheet(SUM_SHEET)
    ws.sheet_properties.tabColor = C_NAVY
    page(ws)
    widths(ws, {"A": 5, "B": 11, "C": 28, "D": 13, **{get_column_letter(c): 11 for c in range(5, 17)},
                "Q": 30})
    title(ws, "A1:Q1", f'="ЗВЕДЕННЯ ППО — "&{q(REF)}!$B$2')
    ws.row_dimensions[1].height = 28
    label(ws, "A3:B3", "Дата звіту")
    inp(ws, "C3", fmt="dd.mm.yyyy", align=CENTER)
    ws["C3"].font = font(12, True, C_NAVY)
    date_dv(ws, ["C3"])
    ws["A4"] = "← вкажіть дату; усе інше рахується автоматично з аркушів підрозділів"
    style(ws, "A4:D4", fnt=font(8, italic=True, color="5B6B7F"), border=None, align=LEFT, merge=True)

    tot = U_LAST + 1
    strike_tot = 45
    tile(ws, "E3:F3", "E4:F5", "Виявлено цілей", f"=E{tot}")
    tile(ws, "G3:H3", "G4:H5", "Застосувань", f"=F{tot}")
    tile(ws, "I3:J3", "I4:J5", "Знищено", f"=M{tot}")
    tile(ws, "K3:L3", "K4:L5", "Ефективність", f"=N{tot}", fmt=PCT)
    tile(ws, "M3:N3", "M4:N5", "Витрачено бортів", f"=O{tot}")
    tile(ws, "O3:P3", "O4:P5", "Ударів противника", f"=SUM(D{strike_tot}:F{strike_tot})")
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 20

    def unit_cols(r, i):
        """№, код, назва, стан аркуша — спільні для трьох таблиць."""
        n = 5 + i
        auto(ws, f"A{r}", f'=IF(B{r}="","",{i + 1})')
        auto(ws, f"B{r}", f'=IF({q(REF)}!$A${n}="","",{q(REF)}!$A${n})', bold=True)
        auto(ws, f"C{r}", f'=IF(B{r}="","",IF({q(REF)}!$B${n}="",B{r},{q(REF)}!$B${n}))', align=LEFT)

    # --- 1. Бойова робота
    section(ws, "A8:Q8", "БОЙОВА РОБОТА ПО ПІДРОЗДІЛАХ")
    for rng, t in [("A9:A10", "№"), ("B9:B10", "Код"), ("C9:C10", "Підрозділ"), ("D9:D10", "Аркуш"),
                   ("E9:E10", "Виявлено цілей"), ("F9:F10", "Застосовано всього"),
                   ("G9:H9", "FPV-перехоплювачі"), ("I9:J9", "Стрілецька зброя"),
                   ("K9:L9", "Інші засоби"), ("M9:M10", "Знищено всього"), ("N9:N10", "Ефектив-ність"),
                   ("O9:O10", "Витрачено бортів"), ("P9:P10", "Втрачено бортів"), ("Q9:Q10", "Черговий")]:
        header(ws, rng, t)
    for col in "GIK":
        header(ws, f"{col}10", "застос.")
    for col in "HJL":
        header(ws, f"{col}10", "знищ.")
    ws.row_dimensions[9].height = 26

    J = lambda c: f"${c}${J_FIRST}:${c}${J_LAST}"
    for i in range(UNIT_SLOTS):
        r = U_FIRST + i
        unit_cols(r, i)
        code = f"$B{r}"
        auto(ws, f"D{r}", f'=IF({code}="","",IF(ISREF({ind(code, "A1")}),"✓","немає аркуша"))')
        ok = f'$D{r}<>"✓"'
        f = {
            "E": f'COUNTIFS({ind(code, J("C"))},"?*",{ind(code, J("C"))},"<>{PATROL}")',
            "F": f'COUNTIFS({ind(code, J("F"))},"?*")',
            "G": f'COUNTIFS({ind(code, J("N"))},"{G_FPV}")',
            "H": f'COUNTIFS({ind(code, J("N"))},"{G_FPV}",{ind(code, J("I"))},"{DESTROYED}")',
            "I": f'COUNTIFS({ind(code, J("N"))},"{G_SMALL}")',
            "J": f'COUNTIFS({ind(code, J("N"))},"{G_SMALL}",{ind(code, J("I"))},"{DESTROYED}")',
            "K": f'COUNTIFS({ind(code, J("N"))},"{G_OTHER}")',
            "L": f'COUNTIFS({ind(code, J("N"))},"{G_OTHER}",{ind(code, J("I"))},"{DESTROYED}")',
            "M": f'COUNTIFS({ind(code, J("I"))},"{DESTROYED}")',
            "O": f'SUM({ind(code, J("J"))})',
            "P": f'SUM({ind(code, J("K"))})',
        }
        for col, expr in f.items():
            auto(ws, f"{col}{r}", f'=IF({ok},"",{expr})', fmt=NUM, bold=col == "M")
        auto(ws, f"N{r}", f'=IF({ok},"",IF(N(F{r})=0,"–",M{r}/F{r}))', fmt=PCT)
        auto(ws, f"Q{r}", f'=IF({ok},"",{ind(code, "I3")}&"")', align=LEFT)
    label(ws, f"A{tot}:D{tot}", "РАЗОМ")
    for col in "EFGHIJKLMOP":
        auto(ws, f"{col}{tot}", f"=SUM({col}{U_FIRST}:{col}{U_LAST})", fmt=NUM, bold=True)
    auto(ws, f"N{tot}", f'=IF(F{tot}=0,"–",M{tot}/F{tot})', fmt=PCT, bold=True)
    auto(ws, f"Q{tot}")
    style(ws, f"A{tot}:Q{tot}", f=fill("DDE3EA"))
    ws.conditional_formatting.add(f"M{U_FIRST}:M{U_LAST}",
                                  DataBarRule(start_type="num", start_value=0, end_type="max",
                                              color="7FB3A8", showValue=True))
    ws.conditional_formatting.add(f"D{U_FIRST}:D{U_LAST}",
                                  CellIsRule(operator="equal", formula=['"немає аркуша"'],
                                             fill=fill("FAD9DC"), font=Font(name=FONT, color="8E1B26")))

    # --- 2. Удари противника
    s_head = 28
    s_first = s_head + 2
    assert s_first + UNIT_SLOTS == strike_tot
    section(ws, f"A{s_head}:Q{s_head}", "УДАРИ ПРОТИВНИКА ПО ПІДРОЗДІЛАХ", color=C_SLATE)
    h = s_head + 1
    for col, t in [("A", "№"), ("B", "Код"), ("C", "Підрозділ"), ("D", "РУ"), ("E", "АУ"), ("F", "УДК"),
                   ("P", "Засобів противника, од."), ("Q", "Втрати о/с")]:
        header(ws, f"{col}{h}", t)
    for k in range(ENEMY_SLOTS):
        col = get_column_letter(7 + k)
        header(ws, f"{col}{h}", f'=IF({q(REF)}!$P${5 + k}="","",{q(REF)}!$P${5 + k})')
    ws.row_dimensions[h].height = 30
    S = lambda c: f"${c}${J_FIRST}:${c}${S_LAST}"
    for i in range(UNIT_SLOTS):
        r = s_first + i
        unit_cols(r, i)
        code, ok = f"$B{r}", f'$D${U_FIRST + i}<>"✓"'
        for col, kind in zip("DEF", ["РУ", "АУ", "УДК"]):
            auto(ws, f"{col}{r}", f'=IF({ok},"",COUNTIFS({ind(code, S("R"))},"{kind}"))', fmt=NUM)
        for k in range(ENEMY_SLOTS):
            col = get_column_letter(7 + k)
            auto(ws, f"{col}{r}", f'=IF(OR({ok},{col}${h}=""),"",COUNTIFS({ind(code, S("S"))},{col}${h}))',
                 fmt=NUM)
        auto(ws, f"P{r}", f'=IF({ok},"",SUM({ind(code, S("T"))}))', fmt=NUM)
        auto(ws, f"Q{r}", f'=IF({ok},"",SUM({ind(code, S("V"))}))', fmt=NUM)
    label(ws, f"A{strike_tot}:C{strike_tot}", "РАЗОМ")
    for c in range(4, 18):
        col = get_column_letter(c)
        auto(ws, f"{col}{strike_tot}", f"=SUM({col}{s_first}:{col}{strike_tot - 1})", fmt=NUM, bold=True)
    style(ws, f"A{strike_tot}:Q{strike_tot}", f=fill("DDE3EA"))

    # --- 3. Сили і засоби
    f_head = strike_tot + 2
    f_first = f_head + 3
    section(ws, f"A{f_head}:Q{f_head}", "СИЛИ І ЗАСОБИ ПО ПІДРОЗДІЛАХ")
    h = f_head + 1
    for rng, t in [(f"A{h}:A{h + 1}", "№"), (f"B{h}:B{h + 1}", "Код"), (f"C{h}:C{h + 1}", "Підрозділ"),
                   (f"D{h}:D{h + 1}", "Зміна"), (f"E{h}:H{h}", "РЛС (наявність)"),
                   (f"I{h}:K{h}", "Робота РЛС"), (f"L{h}:M{h}", "Екіпажі"),
                   (f"N{h}:P{h}", "Борти перехоплювачів"), (f"Q{h}:Q{h + 1}", "Черговий")]:
        header(ws, rng, t)
    for col, t in zip("EFGHIJKLMNOP", ["всього", "справні", "ремонт", "склад", "виявлено цілей",
                                        "включень", "год роботи", "всього", "боєготові", "позиція день",
                                        "позиція ніч", "склад"]):
        header(ws, f"{col}{h + 1}", t)
    ws.row_dimensions[h + 1].height = 26
    src = {"E": f"C{RADAR_TOTAL}", "F": f"D{RADAR_TOTAL}", "G": f"E{RADAR_TOTAL}", "H": f"F{RADAR_TOTAL}",
           "I": f"J{RADAR_TOTAL}", "J": f"K{RADAR_TOTAL}", "K": f"L{RADAR_TOTAL}", "L": f"C{CREW_TOTAL}",
           "M": f"H{CREW_TOTAL}", "N": f"D{CREW_TOTAL}", "O": f"E{CREW_TOTAL}", "P": f"F{CREW_TOTAL}"}
    f_tot = f_first + UNIT_SLOTS
    for i in range(UNIT_SLOTS):
        r = f_first + i
        unit_cols(r, i)
        code, ok = f"$B{r}", f'$D${U_FIRST + i}<>"✓"'
        a, b = ind(code, "C3"), ind(code, "E3")
        auto(ws, f"D{r}", f'=IF({ok},"",IF(AND({a}&""="",{b}&""=""),"",'
                          f'IFERROR({hhmm(a)}&"–"&{hhmm(b)},"")))')
        for col, cell in src.items():
            auto(ws, f"{col}{r}", f'=IF({ok},"",N({ind(code, cell)}))',
                 fmt="[h]:mm" if col == "K" else NUM)
        auto(ws, f"Q{r}", f'=IF({ok},"",{ind(code, "I3")}&"")', align=LEFT)
    label(ws, f"A{f_tot}:D{f_tot}", "РАЗОМ")
    for col in "EFGHIJKLMNOP":
        auto(ws, f"{col}{f_tot}", f"=SUM({col}{f_first}:{col}{f_tot - 1})",
             fmt="[h]:mm" if col == "K" else NUM, bold=True)
    auto(ws, f"Q{f_tot}")
    style(ws, f"A{f_tot}:Q{f_tot}", f=fill("DDE3EA"))

    ws.freeze_panes = "A8"
    protect(ws)
    return ws


# ================================================================ Розрахунок + За типами цілей
CALC = "Розрахунок"
CALC_BLOCKS = ["Виявлено", "Застосовано", "ФПВ знищено", "Стр.зб знищено", "Інші знищено", "Знищено всього"]
BLOCK_W = UNIT_SLOTS + 2         # підрозділи + «Разом» + проміжок
T_FIRST = 5
T_LAST = T_FIRST + TYPE_SLOTS - 1


def calc_col(block, unit=None):
    """Стовпець у «Розрахунку»: unit=None → стовпець «Разом» блоку."""
    start = 2 + block * BLOCK_W
    return get_column_letter(start + (UNIT_SLOTS if unit is None else unit))


def build_calc(wb):
    ws = wb.create_sheet(CALC)
    ws.sheet_properties.tabColor = "8C96A3"
    ws.sheet_view.showGridLines = False
    title(ws, f"A1:{calc_col(len(CALC_BLOCKS) - 1)}1",
          "РОЗРАХУНОК — службовий аркуш (рахується автоматично, не редагується)", size=11)
    ws.column_dimensions["A"].width = 26
    header(ws, "A4", "Тип цілі")
    J = lambda c: f"${c}${J_FIRST}:${c}${J_LAST}"
    extra = [
        "",
        lambda u: f',{ind(u, J("F"))},"?*"',
        lambda u: f',{ind(u, J("N"))},"{G_FPV}",{ind(u, J("I"))},"{DESTROYED}"',
        lambda u: f',{ind(u, J("N"))},"{G_SMALL}",{ind(u, J("I"))},"{DESTROYED}"',
        lambda u: f',{ind(u, J("N"))},"{G_OTHER}",{ind(u, J("I"))},"{DESTROYED}"',
        lambda u: f',{ind(u, J("I"))},"{DESTROYED}"',
    ]
    for t in range(TYPE_SLOTS):
        r = T_FIRST + t
        auto(ws, f"A{r}", f'=IF({q(REF)}!$D${5 + t}="","",{q(REF)}!$D${5 + t})', align=LEFT)
    for b, name in enumerate(CALC_BLOCKS):
        section(ws, f"{calc_col(b, 0)}3:{calc_col(b)}3", name, color=C_SLATE)
        for u in range(UNIT_SLOTS):
            col = calc_col(b, u)
            ws.column_dimensions[col].width = 7
            header(ws, f"{col}4", f'=IF({q(REF)}!$A${5 + u}="","",{q(REF)}!$A${5 + u})')
            code = f"{col}$4"
            ex = extra[b](code) if b else ""
            for t in range(TYPE_SLOTS):
                r = T_FIRST + t
                auto(ws, f"{col}{r}", f'=IF(OR($A{r}="",{code}=""),0,'
                                      f'IFERROR(COUNTIFS({ind(code, J("C"))},$A{r}{ex}),0))', fmt=NUM)
        tc = calc_col(b)
        ws.column_dimensions[tc].width = 8
        header(ws, f"{tc}4", "Разом")
        for t in range(TYPE_SLOTS):
            r = T_FIRST + t
            auto(ws, f"{tc}{r}", f"=SUM({calc_col(b, 0)}{r}:{calc_col(b, UNIT_SLOTS - 1)}{r})",
                 fmt=NUM, bold=True)
        ws.column_dimensions[get_column_letter(2 + b * BLOCK_W + UNIT_SLOTS + 1)].width = 2
    ws.freeze_panes = "B5"
    protect(ws)
    return ws


def build_types(wb):
    ws = wb.create_sheet("За типами цілей")
    ws.sheet_properties.tabColor = C_NAVY
    page(ws)
    widths(ws, {"A": 5, "B": 26, **{get_column_letter(c): 11 for c in range(3, 19)}})
    title(ws, "A1:R1", f'="ЦІЛІ ЗА ТИПАМИ — "&{q(REF)}!$B$2&", "&'
                       f'IF({q(SUM_SHEET)}!$C$3="","дата не вказана",'
                       f'RIGHT("0"&DAY({q(SUM_SHEET)}!$C$3),2)&"."&RIGHT("0"&MONTH({q(SUM_SHEET)}!$C$3),2)&"."&'
                       f'YEAR({q(SUM_SHEET)}!$C$3))')
    ws.row_dimensions[1].height = 28

    section(ws, "A3:I3", "УСЬОГО ЗА ТИПАМИ ЦІЛЕЙ")
    for col, t in zip("ABCDEFGHI", ["№", "Тип цілі", "Виявлено", "Застосовано", "Знищено ФПВ",
                                    "Знищено стр.зб", "Знищено іншими", "Знищено всього", "Ефектив-ність"]):
        header(ws, f"{col}4", t)
    ws.row_dimensions[4].height = 30
    tot = T_LAST + 1
    for t in range(TYPE_SLOTS):
        r = T_FIRST + t
        auto(ws, f"A{r}", f'=IF(B{r}="","",{t + 1})')
        auto(ws, f"B{r}", f"={q(CALC)}!$A{r}", align=LEFT)
        for b, col in enumerate("CDEFGH"):
            auto(ws, f"{col}{r}", f'=IF($B{r}="","",{q(CALC)}!{calc_col(b)}{r})', fmt=NUM, bold=col == "H")
        auto(ws, f"I{r}", f'=IF($B{r}="","",IF(N(D{r})=0,"–",H{r}/D{r}))', fmt=PCT)
    label(ws, f"A{tot}:B{tot}", "РАЗОМ (без патрулювання)")
    auto(ws, f"C{tot}", f'=SUM(C{T_FIRST}:C{T_LAST})-SUMIF($B${T_FIRST}:$B${T_LAST},"{PATROL}",'
                        f'C{T_FIRST}:C{T_LAST})', fmt=NUM, bold=True)
    for col in "DEFGH":
        auto(ws, f"{col}{tot}", f"=SUM({col}{T_FIRST}:{col}{T_LAST})", fmt=NUM, bold=True)
    auto(ws, f"I{tot}", f'=IF(D{tot}=0,"–",H{tot}/D{tot})', fmt=PCT, bold=True)
    style(ws, f"A{tot}:I{tot}", f=fill("DDE3EA"))
    ws.conditional_formatting.add(f"H{T_FIRST}:H{T_LAST}",
                                  DataBarRule(start_type="num", start_value=0, end_type="max",
                                              color="7FB3A8", showValue=True))

    m_head = tot + 3
    m_first = m_head + 2
    section(ws, f"A{m_head}:R{m_head}", "ЗНИЩЕНО: ТИП ЦІЛІ × ПІДРОЗДІЛ")
    h = m_head + 1
    header(ws, f"A{h}", "№")
    header(ws, f"B{h}", "Тип цілі")
    for u in range(UNIT_SLOTS):
        col = get_column_letter(3 + u)
        header(ws, f"{col}{h}", f"={q(CALC)}!{calc_col(0, u)}$4")
    header(ws, f"R{h}", "Разом")
    for t in range(TYPE_SLOTS):
        r, cr = m_first + t, T_FIRST + t
        auto(ws, f"A{r}", f'=IF(B{r}="","",{t + 1})')
        auto(ws, f"B{r}", f"={q(CALC)}!$A{cr}", align=LEFT)
        for u in range(UNIT_SLOTS):
            col = get_column_letter(3 + u)
            auto(ws, f"{col}{r}", f'=IF(OR($B{r}="",{col}${h}=""),"",{q(CALC)}!{calc_col(5, u)}{cr})', fmt=NUM)
        auto(ws, f"R{r}", f'=IF($B{r}="","",SUM(C{r}:Q{r}))', fmt=NUM, bold=True)
    mt = m_first + TYPE_SLOTS
    label(ws, f"A{mt}:B{mt}", "РАЗОМ")
    for c in range(3, 19):
        col = get_column_letter(c)
        auto(ws, f"{col}{mt}", f"=SUM({col}{m_first}:{col}{mt - 1})", fmt=NUM, bold=True)
    style(ws, f"A{mt}:R{mt}", f=fill("DDE3EA"))
    ws.freeze_panes = "C5"
    protect(ws)
    return ws


# ================================================================ Інструкція
def build_help(wb, cfg):
    ws = wb.active
    ws.title = "Інструкція"
    ws.sheet_properties.tabColor = C_NAVY
    page(ws, landscape=False)
    widths(ws, {"A": 3, "B": 24, "C": 90})
    title(ws, "A1:C1", "ДОБОВА ДОПОВІДЬ ППО — як користуватися", size=15)
    ws.row_dimensions[1].height = 32
    r = 3

    def head(text):
        nonlocal r
        section(ws, f"B{r}:C{r}", text)
        r += 1

    def line(a, b, f_a=None):
        nonlocal r
        ws[f"B{r}"], ws[f"C{r}"] = a, b
        style(ws, f"B{r}", f=f_a, fnt=font(10, True, C_SLATE), align=LEFT_TOP, border=None)
        style(ws, f"C{r}", fnt=font(10), align=LEFT_TOP, border=None)
        ws.row_dimensions[r].height = max(16, 15 * (1 + len(b) // 95))
        r += 1

    head("ЩОДНЯ")
    line("1. Дата", "На аркуші «Зведення» вкажіть дату звіту (клітинка C3). Вона підтягнеться на всі аркуші.")
    line("2. Підрозділи", "Кожен підрозділ заповнює лише свій аркуш (вкладки з кодами підрозділів): "
                          "черговий і зміна, РЛС, екіпажі та журнал.")
    line("3. Журнал цілей", "Один рядок = одна подія: виявлення, застосування або патрулювання. "
                            "Тип цілі, засіб і результат обирайте зі списку. Текст доповіді у стовпці M "
                            "формується автоматично; його можна скопіювати.")
    line("4. Удари противника", "Праворуч від журналу: РУ, АУ, УДК із кількістю засобів і втратами.")
    line("5. Зведення", "«Зведення», «За типами цілей» і «Розрахунок» рахуються самі; вручну там нічого "
                        "не вводиться, тож ручних помилок у підсумках не буде.")
    line("6. Новий день", "Збережіть файл як копію з датою в назві, потім очистіть жовті клітинки журналів "
                          "(або почніть з чистого шаблону).")
    r += 1
    head("КОЛЬОРИ")
    line("Жовті клітинки", "заповнюються вручну", f_a=fill(C_INPUT))
    line("Сірі клітинки", "рахуються автоматично, захищені від змін", f_a=fill(C_AUTO))
    line("Рядок «Приклад»", "показує формат заповнення; у підсумках не враховується", f_a=fill(C_EXAMPLE))
    r += 1
    head("ПРАВА ДОСТУПУ")
    line("Зараз", "Усі аркуші захищені від випадкових змін БЕЗ пароля: редагувати можна лише жовті клітинки.")
    line("Пароль підрозділу", "Рецензування → Зняти захист аркуша → Захистити аркуш → задайте окремий пароль "
                              "для кожного аркуша підрозділу й передайте його лише відповідальному.")
    line("Адміністратор", "Так само поставте свій пароль на «Зведення», «За типами цілей», «Довідники», "
                          "«Розрахунок» і на структуру книги (Рецензування → Захистити книгу).")
    line("Важливо", "Пароль аркуша в Excel захищає від помилок, але не є надійним захистом даних. "
                    "Для справжнього розмежування доступу зберігайте файл у OneDrive/SharePoint із правами "
                    "на рівні файлу або ведіть окремий файл для кожного підрозділу.")
    r += 1
    head("ЯК ДОДАТИ ПІДРОЗДІЛ")
    line("1", "Рецензування → Захистити книгу (зняти захист структури).")
    line("2", "Правою кнопкою на вкладці будь-якого підрозділу → Перемістити або копіювати → «Створити копію».")
    line("3", "Перейменуйте вкладку на новий код (наприклад, «П4») і очистіть жовті клітинки.")
    line("4", "На аркуші «Довідники» допишіть код (точно як назва вкладки) і повну назву в перший "
              "вільний рядок. Зведення підхопить підрозділ автоматично (до 15 підрозділів).")
    line("Перевірка", "У «Зведенні» стовпець «Аркуш» показує ✓, якщо код у довіднику збігається з вкладкою.")
    r += 1
    head("ДОВІДНИКИ")
    line("Списки", "Типи цілей, засоби ураження, результати, стан екіпажів, засоби противника й типи РЛС "
                   "редагуються на аркуші «Довідники». Нові значення одразу з'являються у списках і в зведенні "
                   "за типами цілей.")
    line("Групи засобів", "Кожному засобу ураження призначено групу (ФПВ / Стр.зб / Інші); за нею "
                          "рахуються стовпці зведення.")
    ws.protection.sheet = True
    return ws


# ================================================================ main
def build(cfg):
    if len(cfg["units"]) > UNIT_SLOTS:
        raise SystemExit(f"Не більше {UNIT_SLOTS} підрозділів")
    wb = Workbook()
    build_help(wb, cfg)
    build_summary(wb)
    build_types(wb)
    for i, unit in enumerate(cfg["units"], 1):
        build_unit(wb, unit, i, cfg.get("radars", []))
    build_reference(wb, cfg)
    build_calc(wb)
    wb.security = WorkbookProtection(lockStructure=True)
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    return wb


def main():
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    with open(sys.argv[1], encoding="utf-8") as fh:
        cfg = json.load(fh)
    build(cfg).save(sys.argv[2])
    print("Збережено:", sys.argv[2])


if __name__ == "__main__":
    main()
