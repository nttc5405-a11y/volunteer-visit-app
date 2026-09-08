# -*- coding: utf-8 -*-
"""產生「志工訪視電子化系統 操作教學」簡報。"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
SHOTS = os.path.join(HERE, "shots")
OUT = os.path.join(HERE, "志工訪視系統_操作教學.pptx")

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

SW, SH = Inches(13.333), Inches(7.5)

prs = Presentation()
prs.slide_width, prs.slide_height = SW, SH
BLANK = prs.slide_layouts[6]


# ────────────────────────────────────────────────────────
# 基本工具
# ────────────────────────────────────────────────────────
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
    """避免出現 0 或負值的寬高，PowerPoint 會拒絕開啟這種檔案"""
    return max(int(v), int(lo))


def textbox(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(int(x), int(y), clamp(w), clamp(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, item in enumerate(lines):
        text, size, bold, color = (item + (TEXT,))[:4] if len(item) == 3 else item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(6)
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
    """等比縮放置中擺放截圖"""
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
    rect(slide, 0, Inches(1.05), SW, Emu(45000), fill=RGBColor(0xF5, 0x9E, 0x0B), radius=False)
    textbox(slide, Inches(0.6), Inches(0.18), Inches(11.5), Inches(0.5),
            [(title, 28, True, WHITE)])
    if sub:
        textbox(slide, Inches(0.6), Inches(0.66), Inches(11.5), Inches(0.32),
                [(sub, 13, False, RGBColor(0xBF, 0xD2, 0xFF))])


def newslide():
    s = prs.slides.add_slide(BLANK)
    bg = s.background.fill
    bg.solid()
    bg.fore_color.rgb = WHITE
    return s


def bullets(slide, x, y, w, items, size=15, gap=Inches(0.52), color=NAVY):
    """圓點清單，items 為 (主文, 說明or None)"""
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


def steps(slide, x, y, w, items, color=NAVY, size=14):
    """編號步驟"""
    yy = y
    for i, (main, note) in enumerate(items, 1):
        c = rect(slide, x, yy, Inches(0.42), Inches(0.42), fill=color, radius=True)
        c.adjustments[0] = 0.5
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        _set_font(p.add_run(), 14, True, WHITE)
        p.runs[0].text = str(i)
        lines = [(main, size, True, TEXT)]
        if note:
            lines.append((note, 12, False, MUTED))
        textbox(slide, x + Inches(0.62), yy + Emu(20000), w - Inches(0.62),
                Inches(0.42) if not note else Inches(0.7), lines)
        yy += Inches(0.62) if not note else Inches(0.88)
    return yy


def callout(slide, x, y, w, h, title, body, accent=AMBER, bg=RGBColor(0xFF, 0xFB, 0xEB)):
    rect(slide, x, y, w, h, fill=bg, linec=accent, lw=1.5)
    rect(slide, x, y, Emu(60000), h, fill=accent, radius=False)
    lines = [(title, 14, True, accent)]
    for b in body:
        lines.append((b, 12.5, False, TEXT))
    textbox(slide, x + Inches(0.26), y + Inches(0.16), w - Inches(0.5), h - Inches(0.3), lines)


def divider(slide, kicker, title, desc, color):
    rect(slide, 0, 0, SW, SH, fill=color, radius=False)
    textbox(slide, Inches(1.2), Inches(2.55), Inches(11), Inches(0.42),
            [(kicker, 16, True, RGBColor(0xC7, 0xD2, 0xFE))])
    textbox(slide, Inches(1.2), Inches(3.0), Inches(11), Inches(1.0),
            [(title, 46, True, WHITE)])
    textbox(slide, Inches(1.2), Inches(4.1), Inches(10), Inches(0.9),
            [(desc, 16, False, RGBColor(0xDB, 0xE4, 0xFF))])


def table(slide, x, y, w, headers, rows, col_w, row_h=Inches(0.52), head_color=NAVY):
    """簡單表格（用矩形拼，避免 pptx 表格樣式難控）"""
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
            bold = (i == 0)
            col = TEXT if i == 0 else MUTED
            textbox(slide, cx + Inches(0.16), yy + Inches(0.13), col_w[i] - Inches(0.2),
                    row_h - Inches(0.2), [(cell, 12.5, bold, col)])
            cx += col_w[i]
        yy += row_h
    rect(slide, x, y, w, yy - y, fill=None, linec=LINE, radius=False)
    return yy


def footer(slide, text):
    textbox(slide, Inches(0.6), SH - Inches(0.5), Inches(12), Inches(0.3),
            [(text, 11, False, MUTED)])


# ════════════════════════════════════════════════════════
# 1  封面
# ════════════════════════════════════════════════════════
s = newslide()
rect(s, 0, 0, SW, SH, fill=DARK, radius=False)
rect(s, 0, SH - Inches(0.16), SW, Inches(0.16), fill=RGBColor(0xF5, 0x9E, 0x0B), radius=False)
textbox(s, Inches(1.1), Inches(2.2), Inches(11), Inches(0.5),
        [("臺東縣消防局", 18, True, RGBColor(0xF5, 0x9E, 0x0B))])
textbox(s, Inches(1.1), Inches(2.75), Inches(11), Inches(1.3),
        [("志工居家訪視電子化系統", 48, True, WHITE)])
textbox(s, Inches(1.1), Inches(4.05), Inches(11), Inches(0.6),
        [("操作教學　—　依角色分篇", 24, False, RGBColor(0xBF, 0xD2, 0xFF))])
rect(s, Inches(1.1), Inches(4.95), Inches(4.2), Emu(30000), fill=RGBColor(0x33, 0x4E, 0x8C), radius=False)
textbox(s, Inches(1.1), Inches(5.2), Inches(11), Inches(1.0),
        [("志工　│　分隊承辦人　│　大隊承辦人", 16, False, RGBColor(0x9F, 0xB3, 0xE0))])

# ════════════════════════════════════════════════════════
# 2  系統在做什麼
# ════════════════════════════════════════════════════════
s = newslide()
header(s, "這套系統在做什麼", "把紙本訪視表搬到手機上，並讓資料自動變成可用的統計與名單")
bx, by, bw = Inches(0.9), Inches(1.9), Inches(3.5)
labels = [
    ("① 志工用手機填報", "到案家訪視時，用手機一題一題點選\n系統會自動隱藏不需要問的題目", GREEN),
    ("② 自動存入試算表", "送出後立刻寫進 Google 試算表\n並同步到各分隊專屬試算表", BLUE),
    ("③ 承辦人看統計與名單", "儀表板即時彙整，可篩選、可匯出\n直接產生要回訪的家戶清單", PURPLE),
]
for i, (t, d, c) in enumerate(labels):
    x = bx + i * (bw + Inches(0.55))
    rect(s, x, by, bw, Inches(2.5), fill=WHITE, linec=c, lw=2)
    rect(s, x, by, bw, Inches(0.62), fill=c, radius=False)
    textbox(s, x + Inches(0.25), by + Inches(0.16), bw - Inches(0.5), Inches(0.4),
            [(t, 16, True, WHITE)])
    textbox(s, x + Inches(0.25), by + Inches(0.85), bw - Inches(0.5), Inches(1.5),
            [(d, 13, False, TEXT)])
    if i < 2:
        ar = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, int(x + bw + Inches(0.1)),
                                int(by + Inches(1.05)), int(Inches(0.36)), int(Inches(0.4)))
        ar.fill.solid(); ar.fill.fore_color.rgb = LINE
        ar.line.fill.background(); ar.shadow.inherit = False

callout(s, Inches(0.9), Inches(4.75), Inches(11.55), Inches(1.5), "為什麼要換成電子化？",
        ["• 不用再回辦公室key資料，訪視當下送出就完成",
         "• 改善建議會依填答自動勾選，減少漏勾與不一致",
         "• 哪些家戶要回訪、轄區最缺什麼，系統直接算出來，不必再人工翻紙本統計"],
        accent=BLUE, bg=RGBColor(0xEF, 0xF6, 0xFF))
footer(s, "志工訪視電子化系統　操作教學")

# ════════════════════════════════════════════════════════
# 3  三種角色分工
# ════════════════════════════════════════════════════════
s = newslide()
header(s, "三種角色，各做什麼", "系統會依你的「權限角色」自動決定看得到什麼、能做什麼")
cw = [Inches(2.1), Inches(2.5), Inches(3.5), Inches(3.45)]
table(s, Inches(0.9), Inches(1.75), sum(cw, Emu(0)),
      ["角色", "填寫訪視表", "看得到的資料範圍", "主要任務"],
      [["志工", "要（主要工作）", "看不到統計", "到案家訪視、當場填報送出"],
       ["分隊承辦人", "要", "只有自己所屬單位", "督促分隊填報、追蹤待處理家戶"],
       ["大隊承辦人", "不用", "所屬大隊底下所有單位", "督導所轄各分隊的執行成效"],
       ["管理員", "可", "全部單位", "維護題庫、人員與系統設定"]],
      cw, row_h=Inches(0.62))

callout(s, Inches(0.9), Inches(5.15), Inches(11.55), Inches(1.35), "角色是誰設定的？",
        ["由管理員在 Google 試算表的「人員帳號管理」工作表設定「權限角色」欄。",
         "改完之後，該人員下次登入就會套用新的角色，不需要重新安裝或設定任何東西。"],
        accent=NAVY, bg=RGBColor(0xEF, 0xF6, 0xFF))
footer(s, "志工訪視電子化系統　操作教學")

# ════════════════════════════════════════════════════════
# 4  共同：登入
# ════════════════════════════════════════════════════════
s = newslide()
header(s, "第一步：登入（三種角色共通）", "不需要帳號密碼，用身分證與手機末三碼即可")
picture(s, "01_login.png", Inches(0.8), Inches(1.6), Inches(4.2), Inches(5.2))
steps(s, Inches(5.6), Inches(1.9), Inches(7.0),
      [("輸入身分證字號末 3 碼", "例：身分證為 V220031182，就輸入 182"),
       ("輸入手機號碼末 3 碼", "例：手機為 0987-838406，就輸入 406"),
       ("按「進入填寫系統」", "若有同名同碼的人員，會跳出選單請你確認姓名"),
       ("系統依角色帶你到對應畫面", "志工與分隊承辦人 → 填報頁；大隊承辦人 → 儀表板")],
      color=NAVY, size=15)
callout(s, Inches(5.6), Inches(5.15), Inches(7.0), Inches(1.35), "登入不了怎麼辦",
        ["• 先確認末三碼有沒有輸錯（身分證是「字號」的末三碼，不是出生年）",
         "• 仍失敗代表你的資料還沒建到系統裡，請聯絡分隊承辦人協助新增"],
        accent=RED, bg=RGBColor(0xFE, 0xF2, 0xF2))
footer(s, "志工訪視電子化系統　操作教學")

# ════════════════════════════════════════════════════════
# 5  志工篇 分隔
# ════════════════════════════════════════════════════════
s = newslide()
divider(s, "PART 1", "志工篇", "你的工作：到案家訪視，用手機把訪視表填完並送出", GREEN)

# 6  志工流程
s = newslide()
header(s, "志工：一次訪視的完整流程", "整個流程在手機上完成，離開案家前就能送出", GREEN)
steps(s, Inches(0.9), Inches(1.8), Inches(5.6),
      [("登入系統", "身分證＋手機末三碼"),
       ("選擇宣導類型", "防火宣導 或 防災宣導，題目會自動切換"),
       ("填寫案家基本資料", "姓名、地址、屋齡、家庭成員等"),
       ("逐題作答", "點選即可，不需打字")],
      color=GREEN, size=15)
steps(s, Inches(6.9), Inches(1.8), Inches(5.6),
      [("確認改善建議", "系統已自動勾選，可再手動增減"),
       ("預覽檢查", "確認有沒有漏答"),
       ("請受訪者簽名", "手指在螢幕上簽"),
       ("送出", "資料立即進入系統")],
      color=GREEN, size=15)
# 讓右欄編號接續 5~8
callout(s, Inches(0.9), Inches(5.5), Inches(11.55), Inches(1.2), "整份表大約多久？",
        ["熟悉後一戶約 5～8 分鐘。不需要的題目系統會自動隱藏，實際要點的比紙本少。"],
        accent=GREEN, bg=RGBColor(0xEC, 0xFD, 0xF5))
footer(s, "志工篇")

# 7  基本資料
s = newslide()
header(s, "志工：填寫案家基本資料", "帶 ＊ 的欄位一定要填，其餘可略過", GREEN)
picture(s, "02_form_basic.png", Inches(0.8), Inches(1.55), Inches(4.3), Inches(5.3))
bullets(s, Inches(5.7), Inches(1.9), Inches(7.0),
        [("訪視日期預設今天", "如果是補登前幾天的訪視，記得改成實際訪視日"),
         ("訪視宣導類型決定題目", "選錯的話下面題目會整組換掉，請先確認"),
         ("GPS 定位按一下就好", "用來在地圖上標出訪視位置，定位失敗可略過不影響送出"),
         ("家庭成員人數很重要", "長者、行動不便、幼童的人數會用來判斷高風險家戶，請務必詢問填寫")],
        size=15, color=GREEN)
footer(s, "志工篇")

# 8  作答
s = newslide()
header(s, "志工：逐題作答與自動隱藏", "答「未使用」「否」時，不需要問的細項會自動收起來", GREEN)
picture(s, "03_form_qa.png", Inches(0.8), Inches(1.55), Inches(4.3), Inches(5.3))
bullets(s, Inches(5.7), Inches(1.9), Inches(7.0),
        [("直接點選項即可", "不用打字。點過的題目卡片會變成綠色邊框"),
         ("相依題目自動出現或隱藏", "例：第 1 題答「未使用」桶裝瓦斯，後面瓦斯桶相關題目就不會出現"),
         ("日期題會跳出日曆", "例：瓦斯桶下次檢驗日期，直接選日期不用手打"),
         ("上方進度條顯示完成度", "送出前若還有必填未答，系統會提醒你哪幾題")],
        size=15, color=GREEN)
footer(s, "志工篇")

# 9  改善建議
s = newslide()
header(s, "志工：改善建議會自動勾選", "這是最省時間的功能，但仍要看過一遍", GREEN)
picture(s, "04_form_sug.png", Inches(0.8), Inches(1.55), Inches(4.3), Inches(5.3))
bullets(s, Inches(5.7), Inches(1.85), Inches(7.0),
        [("答「否」就自動勾對應建議", "例：瓦斯桶檢驗逾期 → 自動勾「更換合格瓦斯鋼瓶」"),
         ("改回「是」會自動取消", "不必回頭手動清掉"),
         ("你仍可自己加勾或取消", "你手動勾的項目系統不會擅自拿掉")],
        size=15, color=GREEN)
callout(s, Inches(5.7), Inches(4.75), Inches(6.9), Inches(1.7), "防災宣導的了解程度（D17）",
        ["前面 16 題答「否」的數量會自動換算成「了解／尚可／不了解」，",
         "不需要自己判斷，也不必手動選。"],
        accent=GREEN, bg=RGBColor(0xEC, 0xFD, 0xF5))
footer(s, "志工篇")

# 10  預覽送出
s = newslide()
header(s, "志工：預覽、簽名、送出", "送出前最後確認的三個動作", GREEN)
steps(s, Inches(0.9), Inches(1.9), Inches(5.7),
      [("按「下一步：預覽確認」", "系統會先檢查必填題是否都答了"),
       ("逐項核對", "紅字「未作答」代表該題沒填，可按「返回修改」補上"),
       ("請受訪者簽名", "簽名會全螢幕顯示，手機可橫放讓長輩好簽"),
       ("按「確認送出」", "看到成功訊息才算完成")],
      color=GREEN, size=15)
callout(s, Inches(6.9), Inches(1.9), Inches(5.5), Inches(2.0), "返回修改不會弄丟資料",
        ["按「返回修改」回到填報頁時，先前填的內容與勾選都會完整保留，",
         "改完再按一次「下一步」即可。"],
        accent=GREEN, bg=RGBColor(0xEC, 0xFD, 0xF5))
callout(s, Inches(6.9), Inches(4.1), Inches(5.5), Inches(2.35), "送出後才算完成",
        ["填到一半離開或關掉瀏覽器，資料不會自動送出。",
         "務必在案家現場完成送出，並看到「訪視紀錄已成功儲存」的訊息。"],
        accent=RED, bg=RGBColor(0xFE, 0xF2, 0xF2))
footer(s, "志工篇")

# 11  志工常見狀況
s = newslide()
header(s, "志工：常見狀況怎麼處理", "", GREEN)
cw = [Inches(3.6), Inches(7.95)]
table(s, Inches(0.9), Inches(1.7), sum(cw, Emu(0)), ["狀況", "怎麼做"],
      [["GPS 定位失敗", "可以略過，不影響送出。iPhone 請用 Safari 或 Chrome 並允許定位權限"],
       ["現場沒有網路", "先把表填完，移動到有訊號的地方再按送出；不要關掉瀏覽器"],
       ["選錯宣導類型", "在基本資料區改回正確類型，題目會整組換掉，已答的該類型答案保留"],
       ["受訪者不願簽名", "簽名欄可留空仍能送出，但請在備註或口頭向承辦人說明"],
       ["填錯已送出的資料", "無法自行修改，請聯絡分隊承辦人到試算表更正"]],
      cw, row_h=Inches(0.68))
footer(s, "志工篇")

# ════════════════════════════════════════════════════════
# 分隊承辦人篇
# ════════════════════════════════════════════════════════
s = newslide()
divider(s, "PART 2", "分隊承辦人篇", "你除了要訪視，還要掌握本分隊的進度與待處理家戶", BLUE)

s = newslide()
header(s, "分隊承辦人：你比志工多了什麼", "多一個「統計」入口，看得到自己分隊的所有資料", BLUE)
bullets(s, Inches(0.9), Inches(1.9), Inches(11.5),
        [("填報頁右上角多一顆「統計」按鈕", "志工看不到這顆按鈕；點它就進入數據分析儀表板"),
         ("可看到本分隊全部訪視紀錄", "包含案家姓名、地址、電話——這是個資，請勿外流或截圖轉傳"),
         ("分隊篩選被鎖定為自己的單位", "看不到其他分隊的資料，這是系統從後端就限制好的"),
         ("可以匯出名單交給隊員執行", "把待追蹤、無住警器等家戶匯出成 Excel 可開的檔案")],
        size=16, color=BLUE)
callout(s, Inches(0.9), Inches(5.3), Inches(11.55), Inches(1.3), "個資責任",
        ["名單含案家姓名、住址與電話。請只在公務用途使用，不要以通訊軟體轉傳截圖，",
         "列印的紙本用畢請依規定銷毀。"],
        accent=RED, bg=RGBColor(0xFE, 0xF2, 0xF2))
footer(s, "分隊承辦人篇")

s = newslide()
header(s, "分隊承辦人：儀表板總覽", "上方可先用篩選列限定期間與宣導類型", BLUE)
picture(s, "05_dash_top.png", Inches(0.7), Inches(1.5), Inches(11.9), Inches(5.4))
footer(s, "分隊承辦人篇　—　畫面為示範資料")

s = newslide()
header(s, "分隊承辦人：四張待處理卡片", "點任一張卡片，下方就會展開該群家戶的名單", BLUE)
picture(s, "06_dash_cards.png", Inches(0.7), Inches(1.45), Inches(11.9), Inches(2.85))
cw2 = [Inches(2.6), Inches(8.95)]
table(s, Inches(0.9), Inches(4.55), sum(cw2, Emu(0)), ["卡片", "代表什麼、該怎麼用"],
      [["待追蹤訪視", "訪視時志工勾選「需要追蹤訪視」的家戶 → 排入下次訪視行程"],
       ["無住宅警報器", "沒裝住警器的家戶 → 住警器發送或補助的優先對象"],
       ["逃生風險", "通道堆雜物或鐵窗沒有安全出口 → 最容易致命，優先處理"]],
      cw2, row_h=Inches(0.58))
footer(s, "分隊承辦人篇　—　畫面為示範資料")

s = newslide()
header(s, "分隊承辦人：名單與匯出", "這是最實用的功能：直接產生可以交辦的清單", BLUE)
picture(s, "07_dash_list.png", Inches(0.7), Inches(1.5), Inches(11.9), Inches(3.55))
bullets(s, Inches(0.9), Inches(5.35), Inches(11.5),
        [("名單含「原因」欄", "直接告訴你這戶為什麼被列進來，不必再回頭查原始資料"),
         ("按「匯出 CSV」可用 Excel 開啟", "中文不會亂碼，可直接編修後作為分隊工作底冊")],
        size=14, color=BLUE, gap=Inches(0.42))
footer(s, "分隊承辦人篇　—　畫面為示範資料")

s = newslide()
header(s, "分隊承辦人：自訂「高風險家戶」條件", "不同時期關注的重點不一樣，條件可以自己勾", BLUE)
bullets(s, Inches(0.9), Inches(1.85), Inches(11.5),
        [("七個條件可自由勾選", "屋齡 30 年以上、有長者或行動不便者、未裝住警器、逃生風險、瓦斯桶逾期、有幼童、住 3 樓以上"),
         ("可設定「符合幾項」才算高風險", "預設 2 項。若跑出來的高風險家戶太多，就把門檻調高"),
         ("條件一改，卡片數字、名單、地圖顏色同步變動", "可以立刻看出不同標準下的結果差異")],
        size=16, color=BLUE)
callout(s, Inches(0.9), Inches(4.6), Inches(11.55), Inches(1.9), "實務建議",
        ["• 想找獨居長者的火災風險：只勾「有長者或行動不便者」＋「未裝住警器」，門檻設 2 項",
         "• 想做年度總體檢：七項全勾，門檻設 3 項，抓出最需要投入資源的家戶",
         "• 條件設定只影響你畫面上的顯示，不會改到任何訪視資料"],
        accent=BLUE, bg=RGBColor(0xEF, 0xF6, 0xFF))
footer(s, "分隊承辦人篇")

# ════════════════════════════════════════════════════════
# 大隊承辦人篇
# ════════════════════════════════════════════════════════
s = newslide()
divider(s, "PART 3", "大隊承辦人篇", "你不填訪視表，你的工作是督導所轄各分隊的執行成效", PURPLE)

s = newslide()
header(s, "大隊承辦人：你的定位", "只查閱督導，不需要填寫訪視紀錄", PURPLE)
bullets(s, Inches(0.9), Inches(1.9), Inches(11.5),
        [("登入後直接進入儀表板", "不會經過填報頁，也沒有「回到填報」按鈕"),
         ("看得到所屬大隊底下所有單位", "範圍依「人員帳號管理」的所屬大隊欄自動判定"),
         ("分隊下拉可自由切換隊內各單位", "可以只看某一個分隊，也可以看整個大隊合計"),
         ("標題列會顯示你的可檢視範圍", "登入後先看這一行，確認範圍正確再開始看數字")],
        size=16, color=PURPLE)
callout(s, Inches(0.9), Inches(5.3), Inches(11.55), Inches(1.3), "範圍不對怎麼辦",
        ["標題列列出的單位若少了或多了，代表「人員帳號管理」的所屬大隊欄填得不一致",
         "（例如有人寫「台東大隊」有人寫「臺東大隊」）。請聯絡管理員統一名稱。"],
        accent=PURPLE, bg=RGBColor(0xF5, 0xF3, 0xFF))
footer(s, "大隊承辦人篇")

s = newslide()
header(s, "大隊承辦人：督導看什麼", "重點不是件數，而是各分隊的訪視品質與轄區弱點", PURPLE)
picture(s, "08_dash_anly.png", Inches(0.7), Inches(1.45), Inches(11.9), Inches(4.05))
bullets(s, Inches(0.9), Inches(5.7), Inches(11.5),
        [("改善建議排行＝轄區最該加強宣導的項目", "被提最多次的建議，就是下季度宣導的重點"),
         ("分隊比較看的是普及率而非件數", "住警器普及率只計算有回答的家戶，跑越多防災宣導不會被稀釋")],
        size=13.5, color=PURPLE, gap=Inches(0.38))
footer(s, "大隊承辦人篇　—　畫面為示範資料")

s = newslide()
header(s, "大隊承辦人：風險地圖", "一眼看出高風險家戶集中在哪裡", PURPLE)
picture(s, "09_dash_map.png", Inches(0.7), Inches(1.4), Inches(11.9), Inches(3.6))
cw3 = [Inches(2.2), Inches(9.35)]
table(s, Inches(0.9), Inches(5.2), sum(cw3, Emu(0)), ["顏色", "代表"],
      [["紅色（較大）", "高風險家戶——符合你設定的條件項數已達門檻"],
       ["橙色", "有風險項目但未達門檻"],
       ["綠色", "沒有任何風險項目"]],
      cw3, row_h=Inches(0.42))
footer(s, "大隊承辦人篇　—　畫面為示範資料")

s = newslide()
header(s, "大隊承辦人：建議的每月督導流程", "五個步驟，約 15 分鐘", PURPLE)
steps(s, Inches(0.9), Inches(1.9), Inches(11.5),
      [("篩選上個月的期間", "起訖日期設為上月 1 日至月底"),
       ("看「各分隊安全指標比較」", "住警器普及率偏低、或高風險戶數偏多的分隊，列為輔導對象"),
       ("看「改善建議排行」", "前三名就是本大隊最普遍的安全問題，可作為下月宣導主題"),
       ("點「待追蹤訪視」卡片", "確認各分隊是否有積壓未回訪的家戶"),
       ("匯出高風險名單", "於分隊工作會報時交辦，並追蹤處理情形")],
      color=PURPLE, size=15)
footer(s, "大隊承辦人篇")

# ════════════════════════════════════════════════════════
# 附錄
# ════════════════════════════════════════════════════════
s = newslide()
header(s, "附錄：題目要增修怎麼辦（管理員）", "題庫是唯一真相來源，不必改程式")
bullets(s, Inches(0.9), Inches(1.85), Inches(11.5),
        [("所有題目都在「訪視題庫」工作表維護", "新增、修改、停用題目都在這裡，改完志工端重新整理就生效"),
         ("新增題目後點選單「同步題庫欄位」", "系統會自動把欄位補到紀錄表與各分隊試算表"),
         ("停用題目請把「啟用」欄改成 FALSE", "絕對不要刪除整列，否則歷史填答會查不回來"),
         ("答「否」自動勾建議也在題庫設定", "填在「答否建議」欄，不需要工程人員協助")],
        size=16)
callout(s, Inches(0.9), Inches(5.3), Inches(11.55), Inches(1.3), "詳細說明",
        ["「訪視題庫」工作表上已貼有一張完整的圖解說明（含每一欄的填法與常犯錯誤）。",
         "操作前先看那張圖即可，不需要再翻這份簡報。"],
        accent=NAVY, bg=RGBColor(0xEF, 0xF6, 0xFF))
footer(s, "附錄")

s = newslide()
header(s, "重要提醒與求助管道", "")
callout(s, Inches(0.9), Inches(1.7), Inches(5.6), Inches(2.3), "個人資料保護",
        ["• 系統內含案家姓名、住址、電話",
         "• 僅限公務使用，勿以通訊軟體轉傳截圖",
         "• 志工看不到統計資料，這是刻意的設計",
         "• 列印的名單用畢請依規定銷毀"],
        accent=RED, bg=RGBColor(0xFE, 0xF2, 0xF2))
callout(s, Inches(6.85), Inches(1.7), Inches(5.6), Inches(2.3), "資料正確性",
        ["• 送出後無法自行修改，送出前務必確認",
         "• 家庭成員人數會影響風險判定，請確實詢問",
         "• 訪視日期若補登請改成實際訪視當日",
         "• 選錯宣導類型請當場改正"],
        accent=AMBER, bg=RGBColor(0xFF, 0xFB, 0xEB))
callout(s, Inches(0.9), Inches(4.3), Inches(11.55), Inches(2.1), "遇到問題找誰",
        ["志工　→　先找所屬分隊承辦人",
         "分隊承辦人　→　找大隊承辦人或系統管理員",
         "系統本身的異常（打不開、送不出、數字明顯不對）　→　請聯絡系統管理員，並說明你在哪一頁、做了什麼動作、看到什麼訊息"],
        accent=NAVY, bg=RGBColor(0xEF, 0xF6, 0xFF))
footer(s, "志工訪視電子化系統　操作教學")

# 封底
s = newslide()
rect(s, 0, 0, SW, SH, fill=DARK, radius=False)
textbox(s, Inches(1.1), Inches(3.0), Inches(11), Inches(1.0),
        [("開始使用吧", 44, True, WHITE)], align=PP_ALIGN.LEFT)
textbox(s, Inches(1.1), Inches(4.1), Inches(11), Inches(0.6),
        [("有任何操作問題，請依上一頁的求助管道反映", 18, False, RGBColor(0xBF, 0xD2, 0xFF))])
rect(s, 0, SH - Inches(0.16), SW, Inches(0.16), fill=RGBColor(0xF5, 0x9E, 0x0B), radius=False)

prs.save(OUT)
print("saved:", OUT, "slides:", len(prs.slides.__iter__.__self__._sldIdLst))
