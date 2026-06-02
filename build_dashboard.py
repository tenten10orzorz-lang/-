from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.chart.label import DataLabelList
import io

wb = Workbook()

# ── palette ──────────────────────────────────────────────────────────────────
C_NAVY   = "1F3864"
C_BLUE   = "2E75B6"
C_LBLUE  = "D6E4F0"
C_GREEN  = "375623"
C_LGREEN = "E2EFDA"
C_RED    = "843C0C"
C_LRED   = "FCE4D6"
C_AMBER  = "7F6000"
C_LAMBER = "FFEB9C"
C_WHITE  = "FFFFFF"
C_LGRAY  = "F2F2F2"
C_DGRAY  = "595959"
C_ORANGE = "C55A11"

def fill(hex_): return PatternFill("solid", fgColor=hex_)
def font(hex_="000000", sz=10, bold=False, name="Yu Gothic"):
    return Font(color=hex_, size=sz, bold=bold, name=name)
def aln(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
thin = Side(style="thin",   color="BBBBBB")
med  = Side(style="medium", color="888888")
def border(l=thin,r=thin,t=thin,b=thin):
    return Border(left=l, right=r, top=t, bottom=b)
MED_BOX = Border(left=med, right=med, top=med, bottom=med)
THIN_BOX = border()

def rc(ws, row, col, val="", bg=C_WHITE, fg="000000", sz=10, bold=False,
       h="center", v="center", wrap=False, brd=THIN_BOX, num=None):
    c = ws.cell(row=row, column=col, value=val)
    c.fill = fill(bg); c.font = font(fg, sz, bold)
    c.alignment = aln(h, v, wrap); c.border = brd
    if num: c.number_format = num
    return c

def rh(ws, row, h): ws.row_dimensions[row].height = h
def cw(ws, col, w): ws.column_dimensions[get_column_letter(col)].width = w
def mg(ws, r1, c1, r2, c2): ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

def section_title(ws, row, c1, c2, text, bg=C_NAVY):
    mg(ws, row, c1, row, c2)
    c = ws.cell(row=row, column=c1, value=text)
    c.fill = fill(bg); c.font = font(C_WHITE, 11, True)
    c.alignment = aln("left", "center"); c.border = MED_BOX
    rh(ws, row, 22)

def kpi_box(ws, r, c, label, value, unit, status, width_merge=1):
    # status: "good"/"warn"/"bad"/"na"
    bg = {"good": C_LGREEN, "warn": C_LAMBER, "bad": C_LRED, "na": C_LGRAY}[status]
    fg_status = {"good": C_GREEN, "warn": C_AMBER, "bad": C_RED, "na": C_DGRAY}[status]
    icon = {"good": "▲", "warn": "▶", "bad": "▼", "na": "－"}[status]

    if width_merge > 1:
        mg(ws, r, c, r, c + width_merge - 1)
        mg(ws, r+1, c, r+1, c + width_merge - 1)
        mg(ws, r+2, c, r+2, c + width_merge - 1)

    rc(ws, r,   c, label,          bg=bg, fg=C_DGRAY,    sz=9,  bold=False, brd=border(t=med, l=med, r=med))
    rc(ws, r+1, c, f"{value}{unit}",bg=bg, fg=C_NAVY,    sz=16, bold=True,  brd=border(l=med, r=med))
    rc(ws, r+2, c, f"{icon}",       bg=bg, fg=fg_status, sz=11, bold=True,  brd=border(b=med, l=med, r=med))
    rh(ws, r,   16); rh(ws, r+1, 28); rh(ws, r+2, 16)

def progress_bar_formula(ws, r, c, pct_cell, label, bg_empty=C_LGRAY, bg_fill=C_BLUE):
    """Write a visual progress bar using filled cells."""
    pass  # implemented inline below


# ══════════════════════════════════════════════════════════════════════════════
# DUMMY DATA
# ══════════════════════════════════════════════════════════════════════════════
DEALER  = "〇〇商事（仮）"
PERIOD  = "2025年10月〜2026年3月（6ヶ月）"

PRODUCTS = ["ワイパー", "デポジットクリーナー", "V-FORCE/VARTA", "SSオイル", "洗車機液剤", "コーティング"]

# 商品別サマリー: [目標数量, 実績数量, 前年数量]
PROD_DATA = {
    "ワイパー":             [480, 312, 390],
    "デポジットクリーナー":  [360, 198, 280],
    "V-FORCE/VARTA":       [200, 175, 160],
    "SSオイル":             [150, 142, 138],
    "洗車機液剤":           [120, 108, 115],
    "コーティング":         [80,  45,  60],
}

# SS別データ: name, wiper[実績,目標,前年], depot[実績,目標,前年], 戦略タイプ
SS_LIST = [
    {"name": "○○SS（盛岡北）",  "wiper": [52, 60, 55],  "depot": [28, 40, 32], "tier": "B", "strategy": "提案強化"},
    {"name": "○○SS（盛岡南）",  "wiper": [78, 80, 70],  "depot": [45, 50, 40], "tier": "A", "strategy": "維持・深耕"},
    {"name": "○○SS（花巻）",    "wiper": [30, 60, 48],  "depot": [15, 30, 22], "tier": "C", "strategy": "集中テコ入れ"},
    {"name": "○○SS（北上）",    "wiper": [60, 70, 58],  "depot": [38, 40, 35], "tier": "B", "strategy": "提案強化"},
    {"name": "○○SS（一関）",    "wiper": [42, 50, 44],  "depot": [30, 35, 28], "tier": "B", "strategy": "提案強化"},
    {"name": "○○SS（水沢）",    "wiper": [25, 40, 30],  "depot": [12, 25, 18], "tier": "C", "strategy": "集中テコ入れ"},
    {"name": "○○SS（青森東）",  "wiper": [15, 30, 20],  "depot": [8,  20, 12], "tier": "C", "strategy": "集中テコ入れ"},
    {"name": "○○SS（弘前）",    "wiper": [10, 30, 14],  "depot": [22, 20, 18], "tier": "C", "strategy": "集中テコ入れ"},
]

STRATEGY_DETAIL = {
    "維持・深耕":   ("目標達成水準。定期フォローを継続し、新商品提案でさらなる深耕を図る",       "◎ 新商品提案 ◎ 追加SKU展開"),
    "提案強化":     ("目標比70〜90%。重点商品の提案頻度を上げ、POPや試供品を活用した販促を実施","◯ 月1回以上訪問 ◯ POP設置支援"),
    "集中テコ入れ": ("目標比50%未満。要因分析を優先し、代理店同行訪問・特別プロモを検討する",    "✗→◯ 代理店同行訪問 ✗→◯ 特別値引き提案"),
}


# ══════════════════════════════════════════════════════════════════════════════
# Sheet 1: サマリー（表紙）
# ══════════════════════════════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "①サマリー"
ws1.sheet_properties.tabColor = "1F3864"
ws1.sheet_view.showGridLines = False

# column widths
for c, w in [(1,2),(2,18),(3,12),(4,12),(5,12),(6,12),(7,12),(8,12),(9,2)]:
    cw(ws1, c, w)

# ── header band ──
# Row 1: top spacer
mg(ws1, 1, 1, 1, 9)
ws1.cell(1, 1).fill = fill(C_NAVY); rh(ws1, 1, 8)
# Row 2: main title
mg(ws1, 2, 1, 2, 9)
c = ws1.cell(2, 1, "特約店 商品別ダッシュボード")
c.fill = fill(C_NAVY); c.font = font(C_WHITE, 18, True)
c.alignment = aln("left", "center"); rh(ws1, 2, 32)
# Row 3: subtitle
mg(ws1, 3, 1, 3, 9)
ws1.cell(3, 1).fill = fill(C_NAVY); rh(ws1, 3, 10)
# Row 4: dealer / period
mg(ws1, 4, 1, 4, 4)
c4a = ws1.cell(4, 1, f"代理店：{DEALER}")
c4a.fill = fill(C_NAVY); c4a.font = font(C_WHITE, 10); c4a.alignment = aln("left","center")
mg(ws1, 4, 5, 4, 9)
c4b = ws1.cell(4, 5, f"集計期間：{PERIOD}")
c4b.fill = fill(C_NAVY); c4b.font = font(C_WHITE, 10); c4b.alignment = aln("left","center")
rh(ws1, 4, 18)
rh(ws1, 5, 6)  # spacer

# ── KPI row ──
rh(ws1, 6, 14)
mg(ws1, 6, 2, 6, 8)
rc(ws1, 6, 2, "■ 全体 KPI",  bg=C_NAVY, fg=C_WHITE, sz=10, bold=True, h="left", brd=MED_BOX)

kpi_labels = [
    ("全商品 達成率", "65%",  "",   "warn"),
    ("ワイパー 達成率","65%", "",   "warn"),
    ("デポ 達成率",  "55%",   "",   "bad"),
    ("前年比（全体）","104%", "",   "good"),
]
for i, (lbl, val, unit, st) in enumerate(kpi_labels):
    kpi_box(ws1, 7, 2 + i*2, lbl, val, unit, st, width_merge=2)

for r in range(7, 10):
    ws1.cell(r, 1).fill = fill(C_WHITE)
rh(ws1, 10, 10)  # spacer

# ── legend ──
mg(ws1, 11, 2, 11, 8)
rc(ws1, 11, 2, "▲ 達成率80%以上：良好    ▶ 50〜79%：要注意    ▼ 50%未満：要強化対応",
   bg=C_LGRAY, fg=C_DGRAY, sz=9, h="left", brd=THIN_BOX)
rh(ws1, 11, 16); rh(ws1, 12, 8)

# ── product table ──
section_title(ws1, 13, 2, 8, "■ 商品別 目標・実績・達成率・前年比 一覧")
rh(ws1, 14, 6)

hdrs = ["商品", "目標数量", "実績数量", "達成率", "前年数量", "前年比", "判定"]
hbg  = [C_BLUE]*7
for ci, (h, bg) in enumerate(zip(hdrs, hbg), start=2):
    rc(ws1, 15, ci, h, bg=bg, fg=C_WHITE, sz=9, bold=True, brd=THIN_BOX)
rh(ws1, 15, 18)

for ri, (prod, (tgt, act, prev)) in enumerate(PROD_DATA.items(), start=16):
    ach  = act / tgt
    yoy  = act / prev
    st   = "good" if ach >= 0.80 else ("warn" if ach >= 0.50 else "bad")
    icon = {"good": "▲ 良好", "warn": "▶ 注意", "bad": "▼ 要対応"}[st]
    ibg  = {"good": C_LGREEN, "warn": C_LAMBER, "bad": C_LRED}[st]
    rfill= C_LGRAY if ri % 2 == 0 else C_WHITE
    focus= prod in ("ワイパー", "デポジットクリーナー")

    rc(ws1, ri, 2, prod, bg=rfill, fg=C_NAVY if focus else "000000",
       bold=focus, h="left", brd=THIN_BOX)
    rc(ws1, ri, 3, tgt, bg=rfill, brd=THIN_BOX)
    rc(ws1, ri, 4, act, bg=rfill, bold=focus, brd=THIN_BOX)
    rc(ws1, ri, 5, ach, bg=rfill, bold=focus, num="0%", brd=THIN_BOX)
    rc(ws1, ri, 6, prev,bg=rfill, brd=THIN_BOX)
    rc(ws1, ri, 7, yoy, bg=rfill, num="0%", brd=THIN_BOX)
    rc(ws1, ri, 8, icon, bg=ibg, fg={"good":C_GREEN,"warn":C_AMBER,"bad":C_RED}[st],
       bold=True, brd=THIN_BOX)
    rh(ws1, ri, 18)

rh(ws1, 16 + len(PROD_DATA), 8)

# ── chart: 達成率 bar ──
row_chart_start = 16 + len(PROD_DATA) + 2
section_title(ws1, row_chart_start, 2, 8, "■ 商品別 達成率 グラフ")

# build a mini data table for chart (hidden-style)
chart_data_row = row_chart_start + 2
rc(ws1, chart_data_row, 2, "商品", bg=C_LGRAY, sz=8)
rc(ws1, chart_data_row, 3, "達成率", bg=C_LGRAY, sz=8)
rc(ws1, chart_data_row, 4, "目標", bg=C_LGRAY, sz=8)
for i, (prod, (tgt, act, prev)) in enumerate(PROD_DATA.items()):
    r = chart_data_row + 1 + i
    ws1.cell(r, 2, prod)
    ws1.cell(r, 3, act / tgt)
    ws1.cell(r, 4, 1.0)
    ws1.cell(r, 3).number_format = "0%"
    ws1.cell(r, 4).number_format = "0%"

chart = BarChart()
chart.type = "bar"
chart.grouping = "clustered"
chart.title = "商品別 達成率"
chart.y_axis.title = "達成率"
chart.x_axis.title = ""
chart.style = 10
chart.width  = 18
chart.height = 9

cats = Reference(ws1, min_col=2, min_row=chart_data_row+1,
                 max_row=chart_data_row+len(PROD_DATA))
data = Reference(ws1, min_col=3, max_col=4,
                 min_row=chart_data_row, max_row=chart_data_row+len(PROD_DATA))
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.series[0].graphicalProperties.solidFill = "2E75B6"
chart.series[1].graphicalProperties.solidFill = "D9D9D9"
chart.series[1].graphicalProperties.line.solidFill = "AAAAAA"

ws1.add_chart(chart, f"B{row_chart_start + 2}")


# ══════════════════════════════════════════════════════════════════════════════
# Sheet 2: ワイパー詳細
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("②ワイパー詳細")
ws2.sheet_properties.tabColor = "2E75B6"
ws2.sheet_view.showGridLines = False

for c, w in [(1,2),(2,20),(3,11),(4,11),(5,11),(6,11),(7,14),(8,2)]:
    cw(ws2, c, w)

# header
for r in range(1,4):
    mg(ws2, r, 1, r, 8); ws2.cell(r,1).fill = fill(C_BLUE); rh(ws2, r, 8)
c = ws2.cell(2, 1, "ワイパー  SS別 実績・目標・戦略")
c.fill=fill(C_BLUE); c.font=font(C_WHITE,16,True); c.alignment=aln("left","center")
rh(ws2,2,30); rh(ws2,4,8)

# ── SS table ──
section_title(ws2, 5, 2, 7, "■ SS別 当期実績 vs 目標", bg=C_BLUE)
hdrs2 = ["SS名", "実績", "目標", "達成率", "前年", "前年比", "判定"]
for ci, h in enumerate(hdrs2, 2):
    rc(ws2, 6, ci, h, bg=C_NAVY, fg=C_WHITE, sz=9, bold=True)
rh(ws2, 6, 18)

for ri, ss in enumerate(SS_LIST, 7):
    act, tgt, prev = ss["wiper"]
    ach = act / tgt; yoy = act / prev
    st  = "good" if ach>=0.80 else ("warn" if ach>=0.50 else "bad")
    icon= {"good":"▲ 良好","warn":"▶ 注意","bad":"▼ 要対応"}[st]
    ibg = {"good":C_LGREEN,"warn":C_LAMBER,"bad":C_LRED}[st]
    rfill = C_LGRAY if ri%2==0 else C_WHITE
    rc(ws2, ri, 2, ss["name"], bg=rfill, h="left")
    rc(ws2, ri, 3, act,  bg=rfill)
    rc(ws2, ri, 4, tgt,  bg=rfill)
    rc(ws2, ri, 5, ach,  bg=rfill, num="0%", bold=True)
    rc(ws2, ri, 6, prev, bg=rfill)
    rc(ws2, ri, 7, yoy,  bg=rfill, num="0%")
    rc(ws2, ri, 8, icon, bg=ibg,
       fg={"good":C_GREEN,"warn":C_AMBER,"bad":C_RED}[st], bold=True)
    rh(ws2, ri, 20)

r_after = 7 + len(SS_LIST)

# total row
tots = [sum(s["wiper"][i] for s in SS_LIST) for i in range(3)]
ach_t = tots[0]/tots[1]; yoy_t = tots[0]/tots[2]
st_t = "good" if ach_t>=0.80 else ("warn" if ach_t>=0.50 else "bad")
rc(ws2, r_after, 2, "合　計", bg=C_BLUE, fg=C_WHITE, bold=True, h="left")
rc(ws2, r_after, 3, tots[0], bg=C_BLUE, fg=C_WHITE, bold=True)
rc(ws2, r_after, 4, tots[1], bg=C_BLUE, fg=C_WHITE, bold=True)
rc(ws2, r_after, 5, ach_t,  bg=C_BLUE, fg=C_WHITE, bold=True, num="0%")
rc(ws2, r_after, 6, tots[2], bg=C_BLUE, fg=C_WHITE, bold=True)
rc(ws2, r_after, 7, yoy_t,  bg=C_BLUE, fg=C_WHITE, bold=True, num="0%")
rc(ws2, r_after, 8, "", bg=C_BLUE)
rh(ws2, r_after, 22)

# ── chart ──
rh(ws2, r_after+2, 8)
section_title(ws2, r_after+3, 2, 7, "■ SS別 実績 vs 目標 グラフ", bg=C_BLUE)

cd = r_after + 5
rc(ws2, cd, 2, "SS名", bg=C_LGRAY, sz=8)
rc(ws2, cd, 3, "実績", bg=C_LGRAY, sz=8)
rc(ws2, cd, 4, "目標", bg=C_LGRAY, sz=8)
for i, ss in enumerate(SS_LIST):
    r = cd+1+i
    ws2.cell(r, 2, ss["name"].replace("○○SS（","").replace("）",""))
    ws2.cell(r, 3, ss["wiper"][0])
    ws2.cell(r, 4, ss["wiper"][1])

ch2 = BarChart()
ch2.type="col"; ch2.grouping="clustered"
ch2.title="SS別 ワイパー 実績 vs 目標"
ch2.style=10; ch2.width=18; ch2.height=10
cats2 = Reference(ws2, min_col=2, min_row=cd+1, max_row=cd+len(SS_LIST))
data2 = Reference(ws2, min_col=3, max_col=4, min_row=cd, max_row=cd+len(SS_LIST))
ch2.add_data(data2, titles_from_data=True)
ch2.set_categories(cats2)
ch2.series[0].graphicalProperties.solidFill = "2E75B6"
ch2.series[1].graphicalProperties.solidFill = "D9D9D9"
ws2.add_chart(ch2, f"B{r_after+4}")


# ══════════════════════════════════════════════════════════════════════════════
# Sheet 3: デポジットクリーナー詳細
# ══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("③デポジットクリーナー詳細")
ws3.sheet_properties.tabColor = "ED7D31"
ws3.sheet_view.showGridLines = False

for c, w in [(1,2),(2,20),(3,11),(4,11),(5,11),(6,11),(7,14),(8,2)]:
    cw(ws3, c, w)

for r in range(1,4):
    mg(ws3,r,1,r,8); ws3.cell(r,1).fill=fill(C_ORANGE); rh(ws3,r,8)
c=ws3.cell(2,1,"デポジットクリーナー  SS別 実績・目標・戦略")
c.fill=fill(C_ORANGE); c.font=font(C_WHITE,16,True); c.alignment=aln("left","center")
rh(ws3,2,30); rh(ws3,4,8)

section_title(ws3,5,2,7,"■ SS別 当期実績 vs 目標", bg=C_ORANGE)
for ci,h in enumerate(hdrs2,2):
    rc(ws3,6,ci,h,bg=C_NAVY,fg=C_WHITE,sz=9,bold=True)
rh(ws3,6,18)

for ri,ss in enumerate(SS_LIST,7):
    act,tgt,prev=ss["depot"]
    ach=act/tgt; yoy=act/prev
    st="good" if ach>=0.80 else ("warn" if ach>=0.50 else "bad")
    icon={"good":"▲ 良好","warn":"▶ 注意","bad":"▼ 要対応"}[st]
    ibg={"good":C_LGREEN,"warn":C_LAMBER,"bad":C_LRED}[st]
    rfill=C_LGRAY if ri%2==0 else C_WHITE
    rc(ws3,ri,2,ss["name"],bg=rfill,h="left")
    rc(ws3,ri,3,act, bg=rfill)
    rc(ws3,ri,4,tgt, bg=rfill)
    rc(ws3,ri,5,ach, bg=rfill,num="0%",bold=True)
    rc(ws3,ri,6,prev,bg=rfill)
    rc(ws3,ri,7,yoy, bg=rfill,num="0%")
    rc(ws3,ri,8,icon,bg=ibg,fg={"good":C_GREEN,"warn":C_AMBER,"bad":C_RED}[st],bold=True)
    rh(ws3,ri,20)

r_after3=7+len(SS_LIST)
tots3=[sum(s["depot"][i] for s in SS_LIST) for i in range(3)]
ach_t3=tots3[0]/tots3[1]; yoy_t3=tots3[0]/tots3[2]
rc(ws3,r_after3,2,"合　計",bg=C_ORANGE,fg=C_WHITE,bold=True,h="left")
rc(ws3,r_after3,3,tots3[0],bg=C_ORANGE,fg=C_WHITE,bold=True)
rc(ws3,r_after3,4,tots3[1],bg=C_ORANGE,fg=C_WHITE,bold=True)
rc(ws3,r_after3,5,ach_t3,bg=C_ORANGE,fg=C_WHITE,bold=True,num="0%")
rc(ws3,r_after3,6,tots3[2],bg=C_ORANGE,fg=C_WHITE,bold=True)
rc(ws3,r_after3,7,yoy_t3,bg=C_ORANGE,fg=C_WHITE,bold=True,num="0%")
rc(ws3,r_after3,8,"",bg=C_ORANGE)
rh(ws3,r_after3,22)

rh(ws3,r_after3+2,8)
section_title(ws3,r_after3+3,2,7,"■ SS別 実績 vs 目標 グラフ",bg=C_ORANGE)
cd3=r_after3+5
rc(ws3,cd3,2,"SS名",bg=C_LGRAY,sz=8)
rc(ws3,cd3,3,"実績",bg=C_LGRAY,sz=8)
rc(ws3,cd3,4,"目標",bg=C_LGRAY,sz=8)
for i,ss in enumerate(SS_LIST):
    r=cd3+1+i
    ws3.cell(r,2,ss["name"].replace("○○SS（","").replace("）",""))
    ws3.cell(r,3,ss["depot"][0])
    ws3.cell(r,4,ss["depot"][1])

ch3=BarChart()
ch3.type="col"; ch3.grouping="clustered"
ch3.title="SS別 デポジットクリーナー 実績 vs 目標"
ch3.style=10; ch3.width=18; ch3.height=10
cats3=Reference(ws3,min_col=2,min_row=cd3+1,max_row=cd3+len(SS_LIST))
data3=Reference(ws3,min_col=3,max_col=4,min_row=cd3,max_row=cd3+len(SS_LIST))
ch3.add_data(data3,titles_from_data=True)
ch3.set_categories(cats3)
ch3.series[0].graphicalProperties.solidFill=C_ORANGE
ch3.series[1].graphicalProperties.solidFill="D9D9D9"
ws3.add_chart(ch3,f"B{r_after3+4}")


# ══════════════════════════════════════════════════════════════════════════════
# Sheet 4: SS戦略マップ
# ══════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("④SS戦略マップ")
ws4.sheet_properties.tabColor = "375623"
ws4.sheet_view.showGridLines = False

for c, w in [(1,2),(2,22),(3,10),(4,10),(5,14),(6,36),(7,30),(8,2)]:
    cw(ws4, c, w)

for r in range(1,4):
    mg(ws4,r,1,r,8); ws4.cell(r,1).fill=fill(C_GREEN); rh(ws4,r,8)
c=ws4.cell(2,1,"SS別 戦略マップ  ─  アクション計画")
c.fill=fill(C_GREEN); c.font=font(C_WHITE,16,True); c.alignment=aln("left","center")
rh(ws4,2,30); rh(ws4,4,8)

# legend
section_title(ws4,5,2,7,"■ 戦略ランク 凡例",bg=C_GREEN)
legend = [
    ("A：維持・深耕",        C_LGREEN, C_GREEN,  "目標達成圏。関係強化・新商品提案"),
    ("B：提案強化",          C_LAMBER, C_AMBER,  "目標比70〜90%。訪問頻度↑・POP活用"),
    ("C：集中テコ入れ",      C_LRED,   C_RED,    "目標比50%未満。代理店同行・特別提案"),
]
for ri,(lbl,bg,fg,desc) in enumerate(legend,6):
    mg(ws4,ri,2,ri,3); rc(ws4,ri,2,lbl,bg=bg,fg=fg,bold=True,h="left")
    mg(ws4,ri,4,ri,7); rc(ws4,ri,4,desc,bg=bg,fg=fg,h="left")
    rh(ws4,ri,18)
rh(ws4,9,8)

# table header
section_title(ws4,10,2,7,"■ SS別 戦略・アクション一覧",bg=C_GREEN)
hdrs4=["SS名","戦略\nランク","重点商品\n達成率","戦略タイプ","現状評価・取り組み方針","推奨アクション"]
for ci,h in enumerate(hdrs4,2):
    rc(ws4,11,ci,h,bg=C_NAVY,fg=C_WHITE,sz=9,bold=True,wrap=True)
rh(ws4,11,30)

for ri,ss in enumerate(SS_LIST,12):
    tier=ss["tier"]
    bg_tier={"A":C_LGREEN,"B":C_LAMBER,"C":C_LRED}[tier]
    fg_tier={"A":C_GREEN, "B":C_AMBER, "C":C_RED}[tier]
    strat=ss["strategy"]
    desc, act_txt = STRATEGY_DETAIL[strat]

    w_act,w_tgt,_=ss["wiper"]
    d_act,d_tgt,_=ss["depot"]
    avg_ach = (w_act/w_tgt + d_act/d_tgt) / 2
    rfill = C_LGRAY if ri%2==0 else C_WHITE

    rc(ws4,ri,2,ss["name"],  bg=rfill, h="left", bold=True)
    rc(ws4,ri,3,f"{tier}",   bg=bg_tier, fg=fg_tier, bold=True, sz=13)
    rc(ws4,ri,4,avg_ach,     bg=rfill, num="0%", bold=True)
    rc(ws4,ri,5,strat,       bg=bg_tier, fg=fg_tier, bold=True)
    rc(ws4,ri,6,desc,        bg=rfill, h="left", wrap=True)
    rc(ws4,ri,7,act_txt,     bg=rfill, h="left", wrap=True)
    rh(ws4,ri,42)

r_note = 12 + len(SS_LIST) + 1
rh(ws4, r_note-1, 8)
mg(ws4,r_note,2,r_note,7)
rc(ws4,r_note,2,
   "※ 達成率は重点2商品（ワイパー・デポジットクリーナー）の平均値。実データ入力後に自動更新。",
   bg=C_LGRAY, fg=C_DGRAY, sz=9, h="left")
rh(ws4,r_note,16)


# ══════════════════════════════════════════════════════════════════════════════
# save
# ══════════════════════════════════════════════════════════════════════════════
out = "/home/user/-/商品別ダッシュボード_テンプレート.xlsx"
wb.save(out)
print(f"Saved → {out}")
