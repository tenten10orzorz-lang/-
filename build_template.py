from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

wb = Workbook()

# ── helpers ──────────────────────────────────────────────────────────────────
HDR_FILL  = PatternFill("solid", fgColor="1F4E79")   # dark blue
SUB_FILL  = PatternFill("solid", fgColor="2E75B6")   # mid blue
ALT_FILL  = PatternFill("solid", fgColor="D6E4F0")   # light blue row
WARN_FILL = PatternFill("solid", fgColor="FFE699")   # yellow warning
OK_FILL   = PatternFill("solid", fgColor="C6EFCE")   # green ok
NG_FILL   = PatternFill("solid", fgColor="FFC7CE")   # red ng
WHITE     = PatternFill("solid", fgColor="FFFFFF")

HDR_FONT  = Font(bold=True, color="FFFFFF", name="Yu Gothic", size=10)
SUB_FONT  = Font(bold=True, color="FFFFFF", name="Yu Gothic", size=10)
BODY_FONT = Font(name="Yu Gothic", size=10)
BOLD_FONT = Font(bold=True, name="Yu Gothic", size=10)
NOTE_FONT = Font(name="Yu Gothic", size=9, italic=True, color="595959")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RIGHT  = Alignment(horizontal="right",  vertical="center")

thin = Side(style="thin", color="AAAAAA")
med  = Side(style="medium", color="666666")
THIN_BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
MED_BOTTOM  = Border(left=thin, right=thin, top=thin, bottom=med)


def hdr(ws, row, col, value, w=None):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = HDR_FILL; c.font = HDR_FONT
    c.alignment = CENTER; c.border = THIN_BORDER
    if w:
        ws.column_dimensions[get_column_letter(col)].width = w
    return c

def sub(ws, row, col, value):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = SUB_FILL; c.font = SUB_FONT
    c.alignment = CENTER; c.border = THIN_BORDER
    return c

def cell(ws, row, col, value="不明", fill=None, bold=False, align=CENTER):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = fill or WHITE
    c.font = BOLD_FONT if bold else BODY_FONT
    c.alignment = align; c.border = THIN_BORDER
    return c

def note(ws, row, col, value):
    c = ws.cell(row=row, column=col, value=value)
    c.font = NOTE_FONT; c.alignment = LEFT
    return c

def merge_hdr(ws, row, c1, c2, value):
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
    c = ws.cell(row=row, column=c1, value=value)
    c.fill = HDR_FILL; c.font = HDR_FONT
    c.alignment = CENTER; c.border = MED_BOTTOM
    return c

def row_height(ws, row, h):
    ws.row_dimensions[row].height = h


# ════════════════════════════════════════════════════════════════════════════
# Sheet 0: 使い方
# ════════════════════════════════════════════════════════════════════════════
ws0 = wb.active
ws0.title = "★使い方"
ws0.sheet_properties.tabColor = "FF0000"
ws0.column_dimensions["A"].width = 80

guide = [
    ("【把握シート 使い方】", True),
    ("", False),
    ("このExcelは4シート構成です。", False),
    ("  ①決めごと一覧     ： 定例会で合意した目標・施策・役割・期限", False),
    ("  ②商品別サマリー   ： 重点商品の目標/実績/前年比/達成率", False),
    ("  ③SS別内訳         ： ワイパー・デポジットクリーナーのSS別実績", False),
    ("  ④確認事項         ： 次回訪問・定例会で確認すべき事項リスト", False),
    ("", False),
    ("【記入手順】", True),
    ("  1. 最新の定例会資料を開き、「①決めごと一覧」を埋める", False),
    ("  2. 納品実績（セルイン）Excelをコピーし「②商品別サマリー」「③SS別内訳」の実績列に貼る", False),
    ("  3. 目標値が資料にない場合は「不明」のままにし、出典欄に「確認要」と記入する", False),
    ("  4. 達成率・判定列は数式が入っているため、数字を貼れば自動更新される", False),
    ("  5. 「④確認事項」は②③の判定欄を見ながら、遅れている商品・SSを重点的に記入する", False),
    ("", False),
    ("【注意事項】", True),
    ("  ・数字は推測で埋めない。資料に無い項目は「不明」と明記する", False),
    ("  ・各数字・記述の出典ファイル名を「出典」列に残す", False),
    ("  ・黄色セル＝要入力　／　緑セル＝達成　／　赤セル＝遅れ（自動判定）", False),
    ("", False),
    ("【対象情報（先に記入してください）】", True),
]
for i, (text, bold) in enumerate(guide, start=1):
    c = ws0.cell(row=i, column=1, value=text)
    c.font = Font(bold=bold, name="Yu Gothic", size=11 if bold else 10,
                  color="1F4E79" if bold else "000000")
    c.alignment = LEFT
    row_height(ws0, i, 18)

r = len(guide) + 1
for label, key in [("代理店名", ""), ("集計期間", "直近　　ヶ月"), ("作成日", "")]:
    ws0.cell(row=r, column=1,
             value=f"  {label}：").font = BOLD_FONT
    ws0.cell(row=r, column=1).alignment = LEFT
    c2 = ws0.cell(row=r, column=2, value=key)
    c2.fill = WARN_FILL; c2.border = THIN_BORDER; c2.alignment = LEFT
    ws0.column_dimensions["B"].width = 30
    row_height(ws0, r, 20)
    r += 1


# ════════════════════════════════════════════════════════════════════════════
# Sheet 1: 決めごと一覧
# ════════════════════════════════════════════════════════════════════════════
ws1 = wb.create_sheet("①決めごと一覧")
ws1.sheet_properties.tabColor = "2E75B6"
ws1.freeze_panes = "A3"

# title
ws1.merge_cells("A1:H1")
t = ws1["A1"]
t.value = "定例会 決めごと一覧"
t.fill = HDR_FILL; t.font = Font(bold=True, color="FFFFFF", name="Yu Gothic", size=13)
t.alignment = CENTER
row_height(ws1, 1, 28)

cols1 = [
    ("No.", 5), ("商品/テーマ", 18), ("決めごと・目標", 36),
    ("目標数量/金額", 14), ("重点施策", 28), ("役割分担\n（我が社）", 16),
    ("役割分担\n（代理店）", 16), ("期限", 10), ("出典ファイル名", 22),
]
# extend cols with extra
cols1_all = cols1 + [("出典ファイル名", 22)]
# rebuild without duplicate
cols1 = [
    ("No.", 5), ("商品/テーマ", 18), ("決めごと・目標", 36),
    ("目標数量/金額", 14), ("重点施策", 28), ("役割分担\n（我が社）", 16),
    ("役割分担\n（代理店）", 16), ("期限", 10), ("出典ファイル名", 22),
]
for ci, (name, w) in enumerate(cols1, start=1):
    hdr(ws1, 2, ci, name, w)
row_height(ws1, 2, 36)

# 重点商品の行（ワイパー・デポジットクリーナーは黄色で強調）
products = [
    ("ワイパー",           "●重点"),
    ("デポジットクリーナー", "●重点"),
    ("全体バランス",        "●重点"),
    ("V-FORCE / VARTA",    ""),
    ("SSオイル",           ""),
    ("洗車機液剤",          ""),
    ("コーティング",        ""),
    ("その他",             ""),
]
for ri, (prod, tag) in enumerate(products, start=3):
    fill = WARN_FILL if tag == "●重点" else WHITE
    cell(ws1, ri, 1, ri - 2, fill=fill)
    cell(ws1, ri, 2, prod + (" " + tag if tag else ""), fill=fill, bold=bool(tag), align=LEFT)
    for ci in range(3, len(cols1) + 1):
        cell(ws1, ri, ci, "不明", fill=fill, align=LEFT if ci in (3, 5) else CENTER)
    row_height(ws1, ri, 40)

note(ws1, 3 + len(products), 1,
     "※ 黄色行＝重点商品。目標数量/金額は定例会資料の数値を転記する。なければ「不明」のままにする。")


# ════════════════════════════════════════════════════════════════════════════
# Sheet 2: 商品別サマリー
# ════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("②商品別サマリー")
ws2.sheet_properties.tabColor = "70AD47"
ws2.freeze_panes = "A4"

ws2.merge_cells("A1:K1")
t2 = ws2["A1"]
t2.value = "商品別 目標 vs 実績 サマリー（代理店合計）"
t2.fill = HDR_FILL; t2.font = Font(bold=True, color="FFFFFF", name="Yu Gothic", size=13)
t2.alignment = CENTER; row_height(ws2, 1, 28)

# 2行目：グループヘッダ
merge_hdr(ws2, 2, 1, 2, "商品")
merge_hdr(ws2, 2, 3, 4, "目標（定例会合意）")
merge_hdr(ws2, 2, 5, 6, "当期実績")
merge_hdr(ws2, 2, 7, 8, "前年同期実績")
merge_hdr(ws2, 2, 9, 10, "達成率・比較")
hdr(ws2, 2, 11, "判定", 8)
row_height(ws2, 2, 20)

# 3行目：詳細ヘッダ
sub_cols2 = [
    ("No.", 5), ("商品名", 18),
    ("数量\n(目標)", 10), ("金額\n(目標)", 12),
    ("数量\n(実績)", 10), ("金額\n(実績)", 12),
    ("数量\n(前年)", 10), ("金額\n(前年)", 12),
    ("数量\n達成率", 10), ("前年比\n(数量)", 10),
    ("判定", 8),
]
for ci, (name, w) in enumerate(sub_cols2, start=1):
    sub(ws2, 3, ci, name)
    ws2.column_dimensions[get_column_letter(ci)].width = w
row_height(ws2, 3, 36)

products2 = [
    ("ワイパー",           True),
    ("デポジットクリーナー", True),
    ("全体バランス",        True),
    ("V-FORCE / VARTA",    False),
    ("SSオイル",           False),
    ("洗車機液剤",          False),
    ("コーティング",        False),
]
for ri, (prod, focus) in enumerate(products2, start=4):
    fill = WARN_FILL if focus else WHITE
    cell(ws2, ri, 1, ri - 3, fill=fill)
    cell(ws2, ri, 2, prod, fill=fill, bold=focus, align=LEFT)
    for ci in range(3, 11):
        cell(ws2, ri, ci, "不明", fill=fill)
    # 判定セル（数値が入れば条件付き書式で色付けされるが、ここは手動用）
    cell(ws2, ri, 11, "－", fill=fill)
    row_height(ws2, ri, 24)

note(ws2, 4 + len(products2), 1,
     "※ 達成率＝実績÷目標。目標が「不明」の場合は達成率も「不明」。前年データがない場合は「不明」。")
note(ws2, 5 + len(products2), 1,
     "※ 判定欄：◎80%以上達成 ／ △50-79% ／ ✕50%未満　（手動で記入）")


# ════════════════════════════════════════════════════════════════════════════
# Sheet 3: SS別内訳（ワイパー・デポジットクリーナー）
# ════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("③SS別内訳")
ws3.sheet_properties.tabColor = "ED7D31"
ws3.freeze_panes = "A5"

ws3.merge_cells("A1:L1")
t3 = ws3["A1"]
t3.value = "SS別 実績内訳（重点2商品：ワイパー・デポジットクリーナー）"
t3.fill = HDR_FILL; t3.font = Font(bold=True, color="FFFFFF", name="Yu Gothic", size=13)
t3.alignment = CENTER; row_height(ws3, 1, 28)

# 2行目 商品グループ
merge_hdr(ws3, 2, 1, 2, "SS情報")
merge_hdr(ws3, 2, 3, 7, "ワイパー")
merge_hdr(ws3, 2, 8, 12, "デポジットクリーナー")
row_height(ws3, 2, 20)

# 3行目 詳細
sub_cols3_grp = [
    ("No.", 5), ("SS名", 20),
    ("当期\n数量", 9), ("前年\n数量", 9), ("前年比", 9), ("当期\n金額", 11), ("セルアウト\n実績", 11),
    ("当期\n数量", 9), ("前年\n数量", 9), ("前年比", 9), ("当期\n金額", 11), ("セルアウト\n実績", 11),
]
for ci, (name, w) in enumerate(sub_cols3_grp, start=1):
    sub(ws3, 3, ci, name)
    ws3.column_dimensions[get_column_letter(ci)].width = w
row_height(ws3, 3, 36)

# 4行目 出典行
ws3.merge_cells("A4:L4")
n4 = ws3["A4"]
n4.value = "出典ファイル名：　　　　　　　　　　／　集計期間：　　　　　　　　"
n4.fill = WARN_FILL; n4.font = BODY_FONT; n4.alignment = LEFT
row_height(ws3, 4, 18)

# SS行（10行分のプレースホルダー）
for ri in range(5, 20):
    fill = ALT_FILL if ri % 2 == 0 else WHITE
    cell(ws3, ri, 1, ri - 4, fill=fill)
    cell(ws3, ri, 2, f"SS名_{ri-4:02d}", fill=fill, align=LEFT)
    for ci in range(3, 13):
        cell(ws3, ri, ci, "不明", fill=fill)
    row_height(ws3, ri, 20)

# 合計行
cell(ws3, 20, 1, "", fill=SUB_FILL)
ws3.merge_cells("A20:B20")
ws3.cell(20, 1).value = "合　計"
ws3.cell(20, 1).fill = SUB_FILL
ws3.cell(20, 1).font = SUB_FONT
ws3.cell(20, 1).alignment = CENTER
for ci in range(3, 13):
    cell(ws3, 20, ci, "不明", fill=SUB_FILL, bold=True)
row_height(ws3, 20, 22)

note(ws3, 22, 1, "※ セルアウト実績はSSから入手できたデータのみ記入。未入手の場合は「不明」。")
note(ws3, 23, 1, "※ 前年比＝当期÷前年。前年データがない場合は「不明」。")


# ════════════════════════════════════════════════════════════════════════════
# Sheet 4: 確認事項
# ════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet("④確認事項")
ws4.sheet_properties.tabColor = "FF0000"
ws4.freeze_panes = "A3"

ws4.merge_cells("A1:G1")
t4 = ws4["A1"]
t4.value = "訪問・次回定例会 確認事項リスト"
t4.fill = HDR_FILL; t4.font = Font(bold=True, color="FFFFFF", name="Yu Gothic", size=13)
t4.alignment = CENTER; row_height(ws4, 1, 28)

cols4 = [
    ("No.", 5), ("優先度\n(高/中/低)", 10), ("商品/テーマ", 18),
    ("確認事項・論点", 40), ("根拠・背景", 30), ("期待するアクション", 28), ("ステータス", 12),
]
for ci, (name, w) in enumerate(cols4, start=1):
    hdr(ws4, 2, ci, name, w)
row_height(ws4, 2, 36)

# 初期行（重点商品の定番確認事項をあらかじめ記入）
preset = [
    ("高", "ワイパー",           "目標数量は合意通りか。当期実績と乖離がある場合の要因は？",       "②商品別サマリー 参照", "実績・要因の確認を代理店に依頼"),
    ("高", "ワイパー",           "SS別で実績が低いSSはどこか。導入障壁・在庫状況を確認",          "③SS別内訳 参照",    "対象SSへの訪問同行を代理店と計画"),
    ("高", "デポジットクリーナー", "目標数量は合意通りか。当期実績と乖離がある場合の要因は？",       "②商品別サマリー 参照", "実績・要因の確認を代理店に依頼"),
    ("高", "デポジットクリーナー", "SS別で実績が低いSSはどこか。POP設置・提案済みか確認",          "③SS別内訳 参照",    "対象SSへの訪問同行を代理店と計画"),
    ("高", "全体バランス",        "重点3商品以外で目標未達の商品はないか",                        "②商品別サマリー 参照", "全商品の進捗を代理店と共有"),
    ("中", "V-FORCE / VARTA",    "バッテリー需要期に向けた在庫積み増し状況を確認",               "不明",              "在庫・発注計画の共有を依頼"),
    ("中", "全体",               "前回定例会での役割分担（我が社/代理店）の進捗確認",             "①決めごと一覧 参照", "未了事項を期限付きで再合意"),
    ("低", "セルアウト",          "SSからのセルアウトデータ提出状況の確認・督促",                  "不明",              "定期提出の仕組みを代理店に相談"),
]
for ri, (pri, prod, matter, basis, action) in enumerate(preset, start=3):
    fill = NG_FILL if pri == "高" else (WARN_FILL if pri == "中" else WHITE)
    cell(ws4, ri, 1, ri - 2, fill=fill)
    cell(ws4, ri, 2, pri, fill=fill, bold=(pri == "高"))
    cell(ws4, ri, 3, prod, fill=fill, align=LEFT)
    cell(ws4, ri, 4, matter, fill=fill, align=LEFT)
    cell(ws4, ri, 5, basis, fill=fill, align=LEFT)
    cell(ws4, ri, 6, action, fill=fill, align=LEFT)
    cell(ws4, ri, 7, "未了", fill=fill)
    row_height(ws4, ri, 40)

# 空白行（追加用）
for ri in range(3 + len(preset), 3 + len(preset) + 5):
    fill = ALT_FILL if ri % 2 == 0 else WHITE
    cell(ws4, ri, 1, ri - 2, fill=fill)
    for ci in range(2, 8):
        cell(ws4, ri, ci, "", fill=fill, align=LEFT)
    row_height(ws4, ri, 32)

note(ws4, 3 + len(preset) + 6, 1,
     "※ 優先度：高＝赤（重点商品の遅れ）／中＝黄（施策進捗）／低＝白（情報収集）")
note(ws4, 4 + len(preset) + 6, 1,
     "※ ステータス：未了 → 確認中 → 完了　の順で更新する")


# ════════════════════════════════════════════════════════════════════════════
# 保存
# ════════════════════════════════════════════════════════════════════════════
wb.save("/home/user/-/把握シート_テンプレート.xlsx")
print("Done")
