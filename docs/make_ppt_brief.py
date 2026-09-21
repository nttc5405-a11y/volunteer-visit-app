# -*- coding: utf-8 -*-
"""產生「志工居家訪視電子化系統」業務簡報（長官簡報版）。

與 make_ppt.py（操作教學版）的差異：
  教學版＝怎麼操作，對象是使用者；本版＝為什麼做、帶來什麼，對象是長官。
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
# 畫面截圖沿用操作教學版的那一組（已隨專案版控）
REPO = os.path.join("C:\\", "Users", "FireFighter", "Desktop", "AI", "志工訪視電子化",
                    ".claude", "worktrees", "question-bank-new-item-adjustments-cbb9f5")
SHOTS = os.path.join(REPO, "docs", "screenshots")
OUT = os.path.join(HERE, "志工訪視系統_業務簡報.pptx")

FONT = "Microsoft JhengHei"

NAVY   = RGBColor(0x1E, 0x3A, 0x8A)
DARK   = RGBColor(0x0F, 0x1F, 0x45)
TEXT   = RGBColor(0x1E, 0x29, 0x3B)
MUTED  = RGBColor(0x64, 0x74, 0x8B)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT  = RGBColor(0xF1, 0xF5, 0xF9)
LINE   = RGBColor(0xCB, 0xD5, 0xE1)
GREEN  = RGBColor(0x0F, 0x76, 0x6E)
BLUE   = RGBColor(0x1D, 0x4E, 0xD8)
PURPLE = RGBColor(0x6D, 0x28, 0xD9)
AMBER  = RGBColor(0xB4, 0x53, 0x09)
RED    = RGBColor(0xB9, 0x1C, 0x1C)
GOLD   = RGBColor(0xF5, 0x9E, 0x0B)

SW, SH = Inches(13.333), Inches(7.5)
FOOT_Y = Inches(7.0)

prs = Presentation()
prs.slide_width, prs.slide_height = SW, SH
BLANK = prs.slide_layouts[6]


# ─── 基本工具（與教學版同一套視覺語言）────────────────────
def _set_font(run, size, bold=False, color=TEXT):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {})
            rPr.append(e)
        e.set("typeface", FONT)


def clamp(v, lo=Emu(1)):
    """寬高不得為 0 或負值，否則 PowerPoint 會拒絕開啟檔案"""
    return max(int(v), int(lo))


def textbox(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space=6):
    tb = slide.shapes.add_textbox(int(x), int(y), clamp(w), clamp(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, item in enumerate(lines):
        text, size, bold, color = (item + (TEXT,))[:4] if len(item) == 3 else item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(space)
        _set_font(p.add_run(), size, bold, color)
        p.runs[0].text = text
    return tb


def rect(slide, x, y, w, h, fill=None, linec=None, radius=True, lw=1.25):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
        int(x), int(y), clamp(w), clamp(h))
    if radius:
        shp.adjustments[0] = 0.06
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if linec is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = linec
        shp.line.width = Pt(lw)
    shp.shadow.inherit = False
    shp.text_frame.word_wrap = True
    return shp


def picture(slide, name, x, y, max_w, max_h, border=True):
    path = os.path.join(SHOTS, name)
    iw, ih = Image.open(path).size
    scale = min(max_w / iw, max_h / ih)
    w, h = int(iw * scale), int(ih * scale)
    px, py = int(x + (max_w - w) / 2), int(y + (max_h - h) / 2)
    if border:
        rect(slide, px - Emu(20000), py - Emu(20000), w + Emu(40000), h + Emu(40000),
             fill=WHITE, linec=LINE, radius=False)
    slide.shapes.add_picture(path, px, py, clamp(w), clamp(h))


def header(slide, title, sub=None, color=NAVY):
    rect(slide, 0, 0, SW, Inches(1.05), fill=color, radius=False)
    rect(slide, 0, Inches(1.05), SW, Emu(45000), fill=GOLD, radius=False)
    textbox(slide, Inches(0.6), Inches(0.18), Inches(11.8), Inches(0.5), [(title, 28, True, WHITE)])
    if sub:
        textbox(slide, Inches(0.6), Inches(0.66), Inches(11.8), Inches(0.32),
                [(sub, 13, False, RGBColor(0xBF, 0xD2, 0xFF))])


def newslide():
    s = prs.slides.add_slide(BLANK)
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = WHITE
    return s


def bullets(slide, x, y, w, items, size=15, gap=Inches(0.52), color=NAVY):
    yy = y
    for main, note in items:
        dot = rect(slide, x, yy + Emu(48000), Emu(95000), Emu(95000), fill=color, radius=True)
        dot.adjustments[0] = 0.5
        h = Inches(0.34) if not note else Inches(0.62)
        lines = [(main, size, True, TEXT)]
        if note:
            lines.append((note, 12.5, False, MUTED))
        textbox(slide, x + Inches(0.32), yy, w - Inches(0.32), h, lines)
        yy += gap if not note else gap + Inches(0.24)
    return yy


def callout(slide, x, y, w, h, title, body, accent=AMBER, bg=RGBColor(0xFF, 0xFB, 0xEB), size=12.5):
    rect(slide, x, y, w, h, fill=bg, linec=accent, lw=1.5)
    rect(slide, x, y, Emu(60000), h, fill=accent, radius=False)
    lines = [(title, 14, True, accent)]
    for b in body:
        lines.append((b, size, False, TEXT))
    textbox(slide, x + Inches(0.26), y + Inches(0.16), w - Inches(0.5), h - Inches(0.3), lines)


def table(slide, x, y, w, headers, rows, col_w, row_h=Inches(0.52), head_color=NAVY, size=12.5):
    cx = x
    rect(slide, x, y, w, row_h, fill=head_color, radius=False)
    for i, htxt in enumerate(headers):
        textbox(slide, cx + Inches(0.16), y + Inches(0.12), col_w[i] - Inches(0.2),
                row_h - Inches(0.2), [(htxt, 13, True, WHITE)])
        cx += col_w[i]
    yy = y + row_h
    for r, row in enumerate(rows):
        rect(slide, x, yy, w, row_h, fill=(WHITE if r % 2 else LIGHT), radius=False)
        cx = x
        for i, cell in enumerate(row):
            textbox(slide, cx + Inches(0.16), yy + Inches(0.13), col_w[i] - Inches(0.2),
                    row_h - Inches(0.2),
                    [(cell, size, (i == 0), TEXT if i == 0 else MUTED)])
            cx += col_w[i]
        yy += row_h
    rect(slide, x, y, w, yy - y, fill=None, linec=LINE, radius=False)
    return yy


def stat(slide, x, y, w, h, number, unit, label, color=NAVY):
    """數據卡"""
    rect(slide, x, y, w, h, fill=WHITE, linec=LINE, lw=1.5)
    rect(slide, x, y, w, Emu(55000), fill=color, radius=False)
    textbox(slide, x, y + Inches(0.34), w, Inches(0.62),
            [(number, 40, True, color)], align=PP_ALIGN.CENTER)
    textbox(slide, x, y + Inches(1.0), w, Inches(0.26),
            [(unit, 12, True, MUTED)], align=PP_ALIGN.CENTER)
    textbox(slide, x, y + Inches(1.3), w, Inches(0.5),
            [(label, 12.5, False, TEXT)], align=PP_ALIGN.CENTER)


def footer(slide, text):
    textbox(slide, Inches(0.6), FOOT_Y + Inches(0.12), Inches(12), Inches(0.3),
            [(text, 11, False, MUTED)])


# ══════════════════════════════════════════════════════════
# 1  封面
# ══════════════════════════════════════════════════════════
s = newslide()
rect(s, 0, 0, SW, SH, fill=DARK, radius=False)
rect(s, 0, SH - Inches(0.16), SW, Inches(0.16), fill=GOLD, radius=False)
textbox(s, Inches(1.1), Inches(1.95), Inches(11), Inches(0.45), [("臺東縣消防局", 18, True, GOLD)])
textbox(s, Inches(1.1), Inches(2.5), Inches(11), Inches(1.3),
        [("志工居家訪視電子化系統", 46, True, WHITE)])
textbox(s, Inches(1.1), Inches(3.75), Inches(11), Inches(0.6),
        [("業務簡報", 26, False, RGBColor(0xBF, 0xD2, 0xFF))])
rect(s, Inches(1.1), Inches(4.6), Inches(3.8), Emu(30000), fill=RGBColor(0x33, 0x4E, 0x8C), radius=False)
textbox(s, Inches(1.1), Inches(4.85), Inches(11), Inches(1.2),
        [("從紙本填報到風險管理　—　訪視資料即時化、督導可視化", 16, False, RGBColor(0x9F, 0xB3, 0xE0))])

# ══════════════════════════════════════════════════════════
# 2  一頁看懂
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "一頁看懂這套系統", "若只看一頁，請看這三點")
cards = [
    ("做了什麼", "把居家訪視紙本表單全面數位化。\n志工到案家用手機逐題點選，\n離開前當場送出，資料立即入庫。", GREEN),
    ("帶來什麼", "不只產生統計數字，更直接產出\n「哪幾戶要回去處理」的工作名單，\n可匯出交辦分隊執行。", BLUE),
    ("花了多少", "機關端零支出：無軟體授權費、\n無伺服器採購、無委外維護費。\n目前由承辦人以個人帳號及\n自費訂閱之雲端服務維持運作。", PURPLE),
]
bx, bw = Inches(0.9), Inches(3.72)
for i, (t, d, c) in enumerate(cards):
    x = bx + i * (bw + Inches(0.32))
    rect(s, x, Inches(1.75), bw, Inches(2.75), fill=WHITE, linec=c, lw=2)
    rect(s, x, Inches(1.75), bw, Inches(0.62), fill=c, radius=False)
    textbox(s, x + Inches(0.25), Inches(1.91), bw - Inches(0.5), Inches(0.4), [(t, 17, True, WHITE)])
    textbox(s, x + Inches(0.25), Inches(2.6), bw - Inches(0.5), Inches(1.7), [(d, 13.5, False, TEXT)])

callout(s, Inches(0.9), Inches(4.8), Inches(11.55), Inches(1.75), "一句話總結",
        ["過去訪視的成果停在紙本與人工統計，這套系統讓每一次訪視當場變成可用的資料，",
         "並自動篩出高風險家戶與待追蹤對象，使防火宣導從「做了多少」轉為「改善了什麼」。"],
        accent=NAVY, bg=RGBColor(0xEF, 0xF6, 0xFF), size=14)
footer(s, "志工居家訪視電子化系統　業務簡報")

# ══════════════════════════════════════════════════════════
# 3  過去作業的痛點
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "過去作業方式的三個瓶頸", "這是本系統要解決的問題")
pains = [
    ("重複作業", "志工在案家填紙本，回分隊再key一次電腦。\n同一份資料寫兩遍，且入檔常延遲數日。"),
    ("統計耗時", "月報、季報靠人工翻表彙整。\n想交叉分析（例如老屋中有多少戶沒裝住警器）\n幾乎做不到。"),
    ("品質不一", "改善建議由志工憑印象勾選，\n同樣的狀況不同人勾出不同結果，\n資料無法橫向比較。"),
]
bx, bw = Inches(0.9), Inches(3.72)
for i, (t, d) in enumerate(pains):
    x = bx + i * (bw + Inches(0.32))
    rect(s, x, Inches(1.8), bw, Inches(2.9), fill=RGBColor(0xFE, 0xF2, 0xF2), linec=RED, lw=1.5)
    textbox(s, x + Inches(0.28), Inches(2.05), bw - Inches(0.5), Inches(0.45), [(t, 19, True, RED)])
    textbox(s, x + Inches(0.28), Inches(2.62), bw - Inches(0.56), Inches(1.9), [(d, 13, False, TEXT)])

callout(s, Inches(0.9), Inches(5.0), Inches(11.55), Inches(1.5), "共同根源",
        ["資料在紙上，就無法被計算。只要訪視成果不是「一開始就以數位形式產生」，",
         "後面所有的統計、追蹤與督導，都必須靠人力補工，且無法即時。"],
        accent=RED, bg=RGBColor(0xFE, 0xF2, 0xF2), size=13.5)
footer(s, "志工居家訪視電子化系統　業務簡報")

# ══════════════════════════════════════════════════════════
# 4  系統架構
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "系統運作架構", "三個環節，全程無紙本、無重複輸入")
flow = [
    ("志工端", "手機網頁填報", "不需安裝App\n以身分證與手機末三碼登入", GREEN),
    ("資料端", "Google 試算表", "送出即入庫\n並自動同步至各分隊試算表", BLUE),
    ("管理端", "數據分析儀表板", "即時統計、風險名單\n依角色分層檢視", PURPLE),
]
bx, bw = Inches(0.95), Inches(3.5)
for i, (tag, t, d, c) in enumerate(flow):
    x = bx + i * (bw + Inches(0.62))
    rect(s, x, Inches(1.85), bw, Inches(2.6), fill=WHITE, linec=c, lw=2)
    rect(s, x, Inches(1.85), bw, Inches(0.55), fill=c, radius=False)
    textbox(s, x + Inches(0.25), Inches(1.97), bw - Inches(0.5), Inches(0.36), [(tag, 15, True, WHITE)])
    textbox(s, x + Inches(0.25), Inches(2.6), bw - Inches(0.5), Inches(0.45), [(t, 19, True, TEXT)])
    textbox(s, x + Inches(0.25), Inches(3.18), bw - Inches(0.5), Inches(1.1), [(d, 13, False, MUTED)])
    if i < 2:
        ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, int(x + bw + Inches(0.14)),
                                int(Inches(2.95)), int(Inches(0.34)), int(Inches(0.4)))
        ar.fill.solid(); ar.fill.fore_color.rgb = LINE
        ar.line.fill.background(); ar.shadow.inherit = False

callout(s, Inches(0.95), Inches(4.75), Inches(11.45), Inches(1.75), "技術選擇的考量",
        ["以 Google 試算表與 Apps Script 建置，前端採免費靜態網頁代管，",
         "機關端不需採購伺服器、不需資料庫維護、亦無委外年度維護費用；",
         "資料以試算表形式保存，可隨時完整匯出備份，不受特定廠商系統綁定。"],
        accent=NAVY, bg=RGBColor(0xEF, 0xF6, 0xFF), size=13.5)
footer(s, "志工居家訪視電子化系統　業務簡報")

# ══════════════════════════════════════════════════════════
# 5  建置與維運成本
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "建置與維運成本", "本系統自建置至今，未動用機關預算")

rect(s, Inches(0.9), Inches(1.75), Inches(5.6), Inches(3.15), fill=WHITE, linec=GREEN, lw=2)
rect(s, Inches(0.9), Inches(1.75), Inches(5.6), Inches(0.6), fill=GREEN, radius=False)
textbox(s, Inches(1.18), Inches(1.9), Inches(5.1), Inches(0.4), [("機關端支出", 18, True, WHITE)])
yy = Inches(2.6)
for item in ["軟體授權費：無", "伺服器採購與機房：無",
             "系統開發標案：無", "年度委外維護費：無"]:
    d = rect(s, Inches(1.2), yy + Emu(50000), Emu(90000), Emu(90000), fill=GREEN, radius=True)
    d.adjustments[0] = 0.5
    textbox(s, Inches(1.48), yy, Inches(4.8), Inches(0.4), [(item, 14.5, False, TEXT)])
    yy += Inches(0.52)

rect(s, Inches(6.85), Inches(1.75), Inches(5.6), Inches(3.15), fill=WHITE, linec=AMBER, lw=2)
rect(s, Inches(6.85), Inches(1.75), Inches(5.6), Inches(0.6), fill=AMBER, radius=False)
textbox(s, Inches(7.13), Inches(1.9), Inches(5.1), Inches(0.4), [("目前由承辦人自行負擔", 18, True, WHITE)])
yy = Inches(2.6)
for main, note in [("雲端儲存空間訂閱", "系統資料庫與簽名檔存放所需"),
                   ("AI 開發工具訂閱", "系統開發、功能擴充與後續調整")]:
    d = rect(s, Inches(7.15), yy + Emu(50000), Emu(90000), Emu(90000), fill=AMBER, radius=True)
    d.adjustments[0] = 0.5
    textbox(s, Inches(7.43), yy, Inches(4.8), Inches(0.75),
            [(main, 14.5, True, TEXT), (note, 12.5, False, MUTED)])
    yy += Inches(0.95)

callout(s, Inches(0.9), Inches(5.2), Inches(11.55), Inches(1.3), "說明",
        ["本系統之雲端儲存空間與開發工具費用，目前均由承辦人自行訂閱負擔，未向機關請領。",
         "相較於委外開發，除節省開發與年度維護經費外，亦免除招標、驗收與需求變更之行政作業。"],
        accent=NAVY, bg=RGBColor(0xEF, 0xF6, 0xFF), size=13.5)
footer(s, "志工居家訪視電子化系統　業務簡報")

# ══════════════════════════════════════════════════════════
# 5  志工端
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "志工端：手機填報，當場完成", "降低志工負擔是導入成敗的關鍵", GREEN)
picture(s, "03_form_qa.png", Inches(0.8), Inches(1.5), Inches(4.0), Inches(5.0))
bullets(s, Inches(5.4), Inches(1.9), Inches(7.3),
        [("全程點選，不需打字", "大按鈕設計，長輩家中光線不佳也看得清楚"),
         ("不需要的題目自動隱藏", "例如答「未使用桶裝瓦斯」，後續瓦斯相關細項即不再出現"),
         ("受訪者可直接在手機簽名", "簽名檔與紀錄一併存檔，免紙本留存"),
         ("送出前自動檢查必填項目", "避免漏答造成事後補件")],
        size=15, color=GREEN)
callout(s, Inches(5.4), Inches(5.1), Inches(7.3), Inches(1.4), "對志工的實際改變",
        ["訪視完成即結案，不必再回分隊重新輸入電腦。"],
        accent=GREEN, bg=RGBColor(0xEC, 0xFD, 0xF5), size=13.5)
footer(s, "志工居家訪視電子化系統　業務簡報　—　畫面為示範資料")

# ══════════════════════════════════════════════════════════
# 6  題庫規模與智慧化
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "訪視題庫規模與智慧化程度", "以下為目前系統實際運作中的數量")
sw_, sx = Inches(2.62), Inches(0.9)
stat(s, sx + 0 * (sw_ + Inches(0.36)), Inches(1.7), sw_, Inches(1.95), "43", "題", "防火宣導 25 題\n防災宣導 18 題", NAVY)
stat(s, sx + 1 * (sw_ + Inches(0.36)), Inches(1.7), sw_, Inches(1.95), "10", "題", "設有連動條件\n不適用時自動隱藏", GREEN)
stat(s, sx + 2 * (sw_ + Inches(0.36)), Inches(1.7), sw_, Inches(1.95), "21", "條", "改善建議自動勾選規則\n答「否」即自動帶出", BLUE)
stat(s, sx + 3 * (sw_ + Inches(0.36)), Inches(1.7), sw_, Inches(1.95), "21", "個", "分隊已完成建置\n資料自動分流同步", PURPLE)

callout(s, Inches(0.9), Inches(3.95), Inches(11.55), Inches(1.4), "品質一致性",
        ["改善建議不再依賴志工個人判斷：只要某項安全檢查答「否」，系統即依既定規則自動勾選對應建議，",
         "確保同樣的現場狀況，在不同志工手上得到相同的處置建議。"],
        accent=BLUE, bg=RGBColor(0xEF, 0xF6, 0xFF), size=13.5)
callout(s, Inches(0.9), Inches(5.5), Inches(11.55), Inches(1.0), "後續維護不需委外",
        ["題目增修、連動規則調整，承辦人直接在試算表操作即可生效，無需修改程式或另編預算。"],
        accent=AMBER, bg=RGBColor(0xFF, 0xFB, 0xEB), size=13.5)
footer(s, "志工居家訪視電子化系統　業務簡報")

# ══════════════════════════════════════════════════════════
# 7  管理端：儀表板
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "管理端：即時數據分析儀表板", "不必等月報，隨時可查、可篩選期間與單位", BLUE)
picture(s, "05_dash_top.png", Inches(0.7), Inches(1.5), Inches(11.9), Inches(5.35))
footer(s, "志工居家訪視電子化系統　業務簡報　—　畫面為示範資料")

# ══════════════════════════════════════════════════════════
# 8  從統計到行動（核心價值）
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "核心價值：從「統計數字」到「工作名單」", "這是本系統與一般統計報表最大的不同", BLUE)
picture(s, "07_dash_list.png", Inches(0.7), Inches(1.42), Inches(11.9), Inches(3.3))
cw = [Inches(2.75), Inches(8.8)]
table(s, Inches(0.9), Inches(4.9), sum(cw, Emu(0)), ["自動產出的名單", "管理上的用途"],
      [["待追蹤訪視", "志工現場判定需複訪者，直接成為下次訪視排程依據"],
       ["無住宅火災警報器", "住警器補助與發送的優先對象清單"],
       ["高風險家戶", "條件可自訂，名單可匯出交辦分隊列管追蹤"]],
      cw, row_h=Inches(0.5))
footer(s, "志工居家訪視電子化系統　業務簡報　—　畫面為示範資料")

# ══════════════════════════════════════════════════════════
# 9  弱勢與風險分析
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "弱勢家戶辨識與風險分布", "把有限的防火資源，投到最需要的地方", PURPLE)
picture(s, "09_dash_map.png", Inches(0.75), Inches(1.5), Inches(6.9), Inches(4.5))
bullets(s, Inches(8.0), Inches(1.85), Inches(4.7),
        [("辨識獨居長者與行動不便家戶", "自訪視填報的家庭成員結構自動彙整"),
         ("多重條件交叉判定高風險", "屋齡、長者、住警器、逃生動線、瓦斯期限等"),
         ("地圖以顏色標示風險等級", "紅色為高風險家戶，可看出集中區域"),
         ("條件與門檻可依政策調整", "不同時期關注重點不同，隨時可改")],
        size=14, color=PURPLE, gap=Inches(0.5))
callout(s, Inches(8.0), Inches(5.15), Inches(4.7), Inches(1.35), "應用場景",
        ["寒流、颱風前的重點關懷對象，",
         "可在數秒內產出名單。"],
        accent=PURPLE, bg=RGBColor(0xF5, 0xF3, 0xFF), size=13)
footer(s, "志工居家訪視電子化系統　業務簡報　—　畫面為示範資料")

# ══════════════════════════════════════════════════════════
# 10  分層督導
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "分層督導機制", "系統依權限角色自動限定可檢視範圍，無需人工控管")
cw = [Inches(2.15), Inches(3.3), Inches(6.1)]
table(s, Inches(0.9), Inches(1.75), sum(cw, Emu(0)),
      ["角色", "可檢視範圍", "管理職能"],
      [["管理員", "全局所有單位", "系統與題庫維護、全局成效掌握"],
       ["大隊承辦人", "所屬大隊各分隊", "督導所轄分隊執行成效、跨分隊比較"],
       ["分隊承辦人", "僅本分隊", "掌握本隊進度、認領並追蹤待處理家戶"],
       ["志工", "不開放統計資料", "專責訪視填報"]],
      cw, row_h=Inches(0.62), size=13)

callout(s, Inches(0.9), Inches(4.55), Inches(5.6), Inches(1.95), "大隊層級的督導工具",
        ["• 各分隊安全指標橫向比較",
         "• 改善建議項目排行（找出轄區共通弱點）",
         "• 待追蹤案件是否積壓之檢核"],
        accent=PURPLE, bg=RGBColor(0xF5, 0xF3, 0xFF), size=13)
callout(s, Inches(6.85), Inches(4.55), Inches(5.6), Inches(1.95), "比較的是品質而非件數",
        ["住警器普及率等指標，僅以實際問到該題的家戶為分母，",
         "避免因各隊防火、防災宣導比重不同而失真。"],
        accent=BLUE, bg=RGBColor(0xEF, 0xF6, 0xFF), size=13)
footer(s, "志工居家訪視電子化系統　業務簡報")

# ══════════════════════════════════════════════════════════
# 11  個資保護
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "個人資料保護措施", "訪視資料含案家姓名、住址、電話，屬個資保護範圍", RED)
bullets(s, Inches(0.9), Inches(1.8), Inches(11.5),
        [("查詢統計資料必須通過身分驗證", "由後端逐次驗證身分，非僅前端隱藏畫面；未通過者無法取得任何資料"),
         ("依角色限縮資料範圍", "分隊承辦人僅能取得本隊資料，系統於後端即過濾，無法藉由網址繞過"),
         ("志工端不開放統計功能", "志工僅能填報，看不到彙整名單與其他案家資料"),
         ("資料以試算表形式自行保管", "未委由第三方廠商代管，存取權限可自行控管，並可隨時匯出備份")],
        size=15, color=RED, gap=Inches(0.56))

callout(s, Inches(0.9), Inches(5.15), Inches(11.55), Inches(1.35), "使用規範建議",
        ["匯出之名單含個資，建議納入既有公文與檔案管理規定：限公務使用、不得以通訊軟體轉傳，",
         "紙本用畢依規定銷毀。此部分建議由業務單位另訂使用須知，隨系統一併實施。"],
        accent=RED, bg=RGBColor(0xFE, 0xF2, 0xF2), size=13.5)
footer(s, "志工居家訪視電子化系統　業務簡報")

# ══════════════════════════════════════════════════════════
# 12  導入效益
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "導入效益", "分別就三個層級說明")
benefit = [
    ("志工", GREEN, ["現場填報即完成，免回隊重複輸入",
                     "不適用題目自動隱藏，實際填答題數少於紙本",
                     "改善建議自動帶出，減少漏勾與判斷負擔"]),
    ("分隊", BLUE, ["隨時掌握本隊訪視進度與品質",
                    "待處理家戶自動成清單，可直接交辦",
                    "免除人工彙整月報之工時"]),
    ("大隊／本局", PURPLE, ["跨分隊成效可橫向比較，督導有據",
                           "轄區安全弱點量化呈現，作為宣導政策依據",
                           "高風險與弱勢家戶可即時調閱"]),
]
bx, bw = Inches(0.9), Inches(3.72)
for i, (t, c, items) in enumerate(benefit):
    x = bx + i * (bw + Inches(0.32))
    rect(s, x, Inches(1.8), bw, Inches(3.5), fill=WHITE, linec=c, lw=2)
    rect(s, x, Inches(1.8), bw, Inches(0.6), fill=c, radius=False)
    textbox(s, x + Inches(0.25), Inches(1.95), bw - Inches(0.5), Inches(0.4), [(t, 18, True, WHITE)])
    yy = Inches(2.65)
    for it in items:
        d = rect(s, x + Inches(0.28), yy + Emu(45000), Emu(85000), Emu(85000), fill=c, radius=True)
        d.adjustments[0] = 0.5
        textbox(s, x + Inches(0.55), yy, bw - Inches(0.85), Inches(0.75), [(it, 13, False, TEXT)])
        yy += Inches(0.82)

callout(s, Inches(0.9), Inches(5.55), Inches(11.55), Inches(0.95), "尚待累積的資料",
        ["量化效益（節省工時、訪視量成長）須待系統實際運作一段期間後，由儀表板數據呈現。"],
        accent=AMBER, bg=RGBColor(0xFF, 0xFB, 0xEB), size=13.5)
footer(s, "志工居家訪視電子化系統　業務簡報")

# ══════════════════════════════════════════════════════════
# 13  後續建議
# ══════════════════════════════════════════════════════════
s = newslide()
header(s, "後續推動建議", "分為近期執行與中長期可擴充方向")
callout(s, Inches(0.9), Inches(1.75), Inches(5.6), Inches(2.5), "近期執行事項",
        ["1. 各分隊人員角色設定完成並辦理教育訓練",
         "2. 訂定名單使用與個資保護作業須知",
         "3. 試行一至二個月後檢討題目與風險條件",
         "4. 將高風險家戶名單納入分隊工作會報列管"],
        accent=NAVY, bg=RGBColor(0xEF, 0xF6, 0xFF), size=13.5)
callout(s, Inches(6.85), Inches(1.75), Inches(5.6), Inches(2.5), "中長期可擴充",
        ["1. 訪視成果與轄內火災案件勾稽分析",
         "2. 住警器補助對象自動產生與追蹤",
         "3. 年度訪視計畫排程與達成率管理",
         "4. 與既有消防業務系統資料介接"],
        accent=PURPLE, bg=RGBColor(0xF5, 0xF3, 0xFF), size=13.5)

rect(s, Inches(0.9), Inches(4.55), Inches(11.55), Inches(1.95), fill=LIGHT, linec=LINE)
textbox(s, Inches(1.2), Inches(4.78), Inches(11), Inches(0.4), [("建議事項", 16, True, NAVY)])
textbox(s, Inches(1.2), Inches(5.25), Inches(11), Inches(1.1),
        [("本系統已完成建置並可正式使用，建議自下一季起全面導入，", 14, False, TEXT),
         ("並於試行期滿後，依儀表板實際數據檢討訪視重點與資源配置。", 14, False, TEXT)])
footer(s, "志工居家訪視電子化系統　業務簡報")

# ══════════════════════════════════════════════════════════
# 14  封底
# ══════════════════════════════════════════════════════════
s = newslide()
rect(s, 0, 0, SW, SH, fill=DARK, radius=False)
rect(s, 0, SH - Inches(0.16), SW, Inches(0.16), fill=GOLD, radius=False)
textbox(s, Inches(1.1), Inches(2.9), Inches(11), Inches(1.0), [("敬請指導", 44, True, WHITE)])
textbox(s, Inches(1.1), Inches(4.0), Inches(11), Inches(0.6),
        [("臺東縣消防局　志工居家訪視電子化系統", 18, False, RGBColor(0xBF, 0xD2, 0xFF))])

prs.save(OUT)
print("saved:", OUT)
