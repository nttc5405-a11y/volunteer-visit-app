# -*- coding: utf-8 -*-
"""產生「訪視題庫修改說明」提醒圖，供貼入 Google Sheets。"""
import os
from PIL import Image, ImageDraw, ImageFont

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "訪視題庫_修改說明.png")

W, H = 1200, 2400        # 先畫在夠高的畫布上，最後裁切
BG        = (255, 255, 255)
NAVY      = (30, 58, 138)
NAVY_SOFT = (239, 244, 255)
TEXT      = (30, 41, 59)
MUTED     = (100, 116, 139)
LINE      = (203, 213, 225)
GREEN     = (5, 150, 105)
GREEN_BG  = (236, 253, 245)
RED       = (185, 28, 28)
RED_BG    = (254, 242, 242)
AMBER     = (180, 83, 9)
AMBER_BG  = (255, 251, 235)

F = "C:/Windows/Fonts/msjh.ttc"
FB = "C:/Windows/Fonts/msjhbd.ttc"

def f(path, size):
    return ImageFont.truetype(path, size)

f_title   = f(FB, 52)
f_sub     = f(F, 24)
f_h       = f(FB, 32)
f_body    = f(F, 24)
f_bodyb   = f(FB, 24)
f_small   = f(F, 21)
f_smallb  = f(FB, 21)
f_mono    = f(F, 22)

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)

M = 60          # 左右邊界
y = 0

# ── 標題列 ──────────────────────────────────────────────
d.rectangle([0, 0, W, 150], fill=NAVY)
d.text((M, 38), "訪視題庫 修改說明", font=f_title, fill=(255, 255, 255))
d.text((M, 100), "改題目不必找人改程式，全部在這張工作表操作即可", font=f_sub, fill=(191, 210, 255))
y = 190


def section(title, color=NAVY):
    """區塊標題"""
    global y
    d.rectangle([M, y, M + 8, y + 34], fill=color)
    d.text((M + 22, y - 2), title, font=f_h, fill=color)
    y += 56


def line(text, font=f_body, fill=TEXT, indent=0, gap=36):
    global y
    d.text((M + indent, y), text, font=font, fill=fill)
    y += gap


def box(lines, bg, border, pad=18, font=f_small, fill=TEXT):
    """有底色的提示框"""
    global y
    h = pad * 2 + len(lines) * 32 - 6
    d.rounded_rectangle([M, y, W - M, y + h], radius=10, fill=bg, outline=border, width=2)
    yy = y + pad
    for t, bold in lines:
        d.text((M + pad + 4, yy), t, font=(f_smallb if bold else font), fill=fill)
        yy += 32
    y += h + 26


def menu_item(name, desc):
    """選單項目：藍底標籤 + 說明"""
    global y
    tw = d.textlength(name, font=f_smallb)
    d.rounded_rectangle([M + 8, y - 4, M + 8 + tw + 28, y + 32], radius=6,
                        fill=NAVY_SOFT, outline=NAVY, width=1)
    d.text((M + 22, y + 1), name, font=f_smallb, fill=NAVY)
    d.text((M + 8 + tw + 46, y + 1), desc, font=f_small, fill=TEXT)
    y += 48


# ── 1. 新增一題 ────────────────────────────────────────
section("① 新增一題")
line("1. 在下方新增一列，「題目代碼」沿用下一個號碼（例：F26、D19）", indent=8)
line("2. 想讓它出現在某題後面，直接把整列拖曳到那個位置", indent=8)
line("3. 上方選單「志工訪視系統 → 同步題庫欄位」點一下", indent=8)
line("4. 完成。志工端重新整理表單就會看到", indent=8, gap=44)

box([("為什麼可以隨便拖曳順序？", True),
     ("答案是照「題目代碼」對應欄位存的，不是照列的先後順序，", False),
     ("所以調整順序、在中間插入新題目，都不會讓舊資料跑位。", False)],
    GREEN_BG, GREEN)

# ── 2. 停用一題 ────────────────────────────────────────
section("② 不想用某一題")
box([("把該列的「啟用」欄改成 FALSE  ——  不要刪除整列！", True),
     ("改成 FALSE：題目不再出現於表單，但欄位與歷年填答完整保留。", False),
     ("直接刪列：欄位對應消失，過去所有訪視紀錄的那一題再也查不回來。", False)],
    RED_BG, RED)

# ── 3. 欄位說明 ────────────────────────────────────────
section("③ 每一欄要填什麼")

cols = [
    ("題目代碼", "同時是紀錄表的欄位名稱。用過就不要改名或重複使用"),
    ("訪視類型", "防火宣導 或 防災宣導"),
    ("依賴條件", "例 F01=是 → 該題答「是」時本題才出現（留空＝一定出現）"),
    ("題目分類", "分類名稱。填「改善建議」的複選題會成為自動勾選的對象"),
    ("題目內容", "顯示給志工看的題目文字"),
    ("題　　型", "是否題／單選題／複選題／簡答題／日期題／數字題"),
    ("選項內容", "用半形逗號分隔，例：是,否。簡答、日期、數字題留空"),
    ("必　　填", "TRUE＝必填，FALSE 或留空＝選填"),
    ("啟　　用", "FALSE＝停用，留空或 TRUE＝啟用"),
    ("答否建議", "該題答「否」時自動勾選的改善建議，多項用半形逗號分隔"),
]

row_h = 46
tbl_top = y
d.rounded_rectangle([M, y, W - M, y + row_h * len(cols)], radius=10, outline=LINE, width=2)
for i, (name, desc) in enumerate(cols):
    ry = tbl_top + i * row_h
    if i % 2 == 0:
        d.rectangle([M + 2, ry + 2, W - M - 2, ry + row_h - 2], fill=(248, 250, 252))
    d.text((M + 20, ry + 11), name, font=f_smallb, fill=NAVY)
    d.text((M + 190, ry + 11), desc, font=f_small, fill=TEXT)
    if i:
        d.line([M + 2, ry, W - M - 2, ry], fill=LINE, width=1)
y = tbl_top + row_h * len(cols) + 34

# ── 4. 自動勾選改善建議 ────────────────────────────────
section("④ 讓某題答「否」時自動勾選改善建議")
line("在該題的「答否建議」欄填入建議文字即可，例如：", indent=8)
y += 4
d.rounded_rectangle([M + 8, y, W - M, y + 78], radius=8, fill=NAVY_SOFT, outline=LINE, width=1)
d.text((M + 28, y + 12), "F16  →  勿在逃生通道堆放雜物", font=f_mono, fill=NAVY)
d.text((M + 28, y + 44), "F02  →  更換合格瓦斯鋼瓶,定期檢查瓦斯桶與皮管", font=f_mono, fill=NAVY)
y += 100

box([("建議文字必須和 F23／D18 的選項一字不差，多一個空格都不算。", True),
     ("不用指定要勾到哪一題，系統會自動找同類型的「改善建議」複選題。", False),
     ("多題對到同一項建議時，要所有觸發題都改成「是」才會自動取消。", False)],
    AMBER_BG, AMBER)

# ── 5. 上方選單 ────────────────────────────────────────
section("⑤ 上方選單「志工訪視系統」")
menu_item("同步題庫欄位", "新增題目後點一下，把欄位補到紀錄表與各分隊試算表")
menu_item("填入預設連動規則", "一鍵帶入預設的「答否建議」（只填空白，不覆蓋你改過的）")
y += 6
box([("選單沒出現？重新整理一次試算表頁面即可。", False),
     ("改完題庫不需要重新部署，志工端下次載入表單就生效。", False)],
    NAVY_SOFT, LINE)

# ── 頁尾 ────────────────────────────────────────────────
y += 14
d.line([M, y, W - M, y], fill=LINE, width=2)
y += 22
d.text((M, y), "半形逗號是「,」不是「，」　｜　題目代碼用過就不要改名或重複使用", font=f_small, fill=MUTED)
y += 32
d.text((M, y), "停用題目請用「啟用」欄改 FALSE，不要刪除整列", font=f_small, fill=MUTED)
y += 56

img = img.crop((0, 0, W, y))
img.save(OUT)
print("OK:", OUT, img.size)
