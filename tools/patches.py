# -*- coding: utf-8 -*-
"""依「原 PPT 頁碼」做的內容調整。build.py 讀完每一頁後會呼叫 apply()。

PPT 改版後若頁序變動，這裡的頁碼要跟著改；找不到指定文字時 build 會停下來並指出是哪一頁。
"""
import copy
import re

GREY, DARK, NAVY = "#3E4756", "#5F6673", "#15233A"

# 整頁刪除
DROP = {18, 69, 73, 108, 111, 115}
# 2026-10-13：第二部份「AI倫理、資安與法規」整段拿掉（原第 25–59 頁，含演練 36/52、參考答案 37/52、手工頁 32/46）
DROP |= set(range(25, 60))

# 插入的其他簡報：(接在主簡報第幾頁之後, 檔名（與主簡報同資料夾）, 頁碼基數)；該簡報第 j 頁在這裡的頁碼是 基數+j
INSERTS = [
    (22, "生成式AI與智慧營運人才培訓授課簡報完成版本(AI 三功能的本質區分).pptx", 200),  # 201：生成式 AI 三大功能
    (79, "生成式AI與智慧營運人才培訓授課簡報完成版本(ChatGPT).pptx", 300),  # 301–312：第三部份 ChatGPT 實作
    (109, "生成式AI與智慧營運人才培訓授課簡報完成版本(Gemini).pptx", 400),  # 401–440：第六部份 Gemini
    (128, "Claude功能與使用全攻略.pptx", 500),  # 501–540：第七部份 Claude（540 是重複的感謝頁，刪）
    (536, "Claude新增加內容至網頁簡報.pptx", 600),  # 601–607：第七部份單元六 Claude Skills（接在 STEP 4–5 之後、使用注意事項之前）
]
DROP |= {128, 540}  # 128：Skills 結語頁（依 2026-10-13 指示拿掉）

# 背景色覆寫（把插入簡報的章名頁改成主簡報的深藍章名頁）
BG = {}

# 整份插入簡報的字色對應（藍色系 → 主簡報的深藍／橘）；色塊與表格的顏色在 restyle() 處理
BLUE_MAP = {
    "#4967E9": "#C2610F", "#D2D9F9": "#EFE3D3", "#BBC6F7": "#EFE3D3", "#F9F9FF": "#FCFBF8",
    "#404155": "#15233A", "#000000": "#15233A", "#1B1B27": "#15233A", "#002060": "#15233A",
    "#910D0D": "#A3500C", "#C00000": "#C2610F", "#F44444": "#C2610F", "#FF0000": "#C2610F",
}
DECK_RECOLOR = {300: BLUE_MAP, 400: BLUE_MAP, 600: BLUE_MAP}


def reorder(keys):
    """第五部份：先講 Skills（原 110–127）再講 Work（原 96–109）"""
    skills = [k for k in keys if 110 <= k <= 127]
    work = [k for k in keys if 96 <= k <= 109]
    out = [k for k in keys if k not in skills and k not in work]
    pos = out.index(95) + 1
    out[pos:pos] = skills + work
    # 三個市場調查演練（原 86–88）搬到第三部份「運用 ChatGPT 進行市場調查分析」章名頁之後
    mk = [k for k in out if k in (86, 87, 88)]
    out = [k for k in out if k not in mk]
    pos = out.index(306) + 1
    out[pos:pos] = mk
    return out


def restyle(els, cmap):
    """色塊、框線、表格儲存格的顏色對應（字色由 build 讀檔時處理）"""
    up = {k.upper(): v for k, v in cmap.items()}
    for e in els:
        for k in ("fill", "stroke"):
            if e.get(k):
                e[k] = up.get(e[k].upper(), e[k])
        for row in e.get("rows", []):
            for cell in row["cells"]:
                cell["css"] = re.sub(r"#[0-9A-Fa-f]{6}", lambda m: up.get(m.group(0).upper(), m.group(0)), cell["css"])


def auto_roles(els, h):
    """沒有標題版面配置的插入頁：上方字最大的文字框當標題，標題上方的當頁首小標；整頁對齊到標題 y≈129"""
    texts = [e for e in els if "paras" in e and not e.get("fill")]
    cands = [e for e in texts if e["y"] < 170] or texts
    zmax = lambda e: max(r["z"] for p in e["paras"] for r in p["runs"])
    t = max(cands, key=zmax)
    t["role"] = "title"
    for e in texts:
        if e is not t and e["y"] + e["h"] <= t["y"] + 2 and e["role"] != "source":
            e["role"] = "eyebrow"
    dy = 129 - t["y"]
    for e in els:
        if e["role"] != "source":
            e["y"] += dy
    # 左右邊界對齊主簡報的版心（85.3–1194.4）：1600 寬的簡報換算後偏左
    body = [e for e in els if e["role"] not in ("eyebrow", "source")]
    x0, x1 = min(e["x"] for e in body), max(e["x"] + e["w"] for e in body)
    if x1 - x0 > 600 and (abs(x0 - 85.3) > 3 or abs(x1 - 1194.4) > 3):
        f = (1194.4 - 85.3) / (x1 - x0)
        for e in body:
            if "paras" not in e and abs(e["w"] - e["h"]) <= 2:  # 圓形只移動，不變形
                e["x"] = 85.3 + (e["x"] + e["w"] / 2 - x0) * f - e["w"] / 2
            else:
                e["x"] = 85.3 + (e["x"] - x0) * f
                e["w"] *= f
                if "cols" in e:
                    e["cols"] = [c * f for c in e["cols"]]
    return t


def bullets(e, items, gap=6, char="•", ind=20):
    """文字框改成圓點條列，沿用原本的字級與顏色"""
    base = e["paras"][0]
    e["paras"], e["free"] = [], True
    for k, item in enumerate(items):
        para = copy.deepcopy(base)
        last = dict(para["runs"][-1] if para["runs"][-1]["text"].strip() else para["runs"][0])
        last["text"] = item
        para["runs"] = [last]
        para["sb"], para["hang"], para["bu"] = (gap if k else 0), None, (char, ind)
        e["paras"].append(para)


def runs_of(e):
    return [r for para in e["paras"] for r in para["runs"]]


def mk_text(like, x, y, w, h, lines, z, color, serif=False, b=True, lh=None):
    """用現有文字元素當樣板，造一個新的文字框（lines：每行一段）"""
    e = copy.deepcopy(like)
    e.update(x=x, y=y, w=w, h=h, role="", fill=None, stroke=None, sw=0, prst="rect", anchor="t", pad=[0, 0, 0, 0])
    e.pop("boost", None)
    e["paras"] = e["paras"][:1]
    for i, line in enumerate(lines):
        if i:
            e["paras"].append(copy.deepcopy(e["paras"][0]))
        para = e["paras"][i]
        para.update(algn=None, sb=0, bu=None, x=None, hang=None, lh=lh or round(z * 1.3, 1))
        para["runs"] = [dict(para["runs"][0], text=line, z=z, color=color, serif=serif, b=b, i=False, u=False, spc=0, href=None)]
    return e


def divider(els, eyebrow, title, desc):
    """整頁改成主簡報的深藍章名頁（版面照原第 95 頁）；呼叫端要把 BG 設成 NAVY"""
    like = next(e for e in els if "paras" in e)
    est = sum(1 if ord(ch) > 0x2E7F else 0.56 for ch in title) * 64  # 標題太長就縮小字級，維持一行
    tz = 64 if est <= 1120 else round(64 * 1120 / est, 1)
    bar = {"tag": "sp", "role": "", "x": 85.3, "y": 248, "w": 106.7, "h": 5.3, "ph": None, "prst": "rect", "adj": 16667,
           "fill": AMBER, "stroke": None, "sw": 0, "arrow": False}
    els[:] = [bar,
              mk_text(like, 85.3, 280, 1142.7, 34, [eyebrow], 22.7, AMBER),
              mk_text(like, 85.3, 338, 1142.7, 76, [title], tz, CREAM, serif=True, lh=80)]
    if desc:
        els.append(mk_text(like, 85.3, 438, 1142.7, 36 * len(desc), desc, 22.7, "#C7D0DC", b=False, lh=36))


# 插入簡報裡的章名頁／單元頁（原本是左右各半的圖＋標題）：改成深藍章名頁
DIVIDERS = {
    601: ("第七部份｜單元六", "Claude Skills：雇用個人專屬的 AI 助理",
          ["把每次都要重複交代的指令，變成可自動觸發、可重複呼叫、可分享的標準作業流程",
           "SKILL.md：個人專屬 AI 助理的「聘書」"]),
    301: ("第三部份", "運用 ChatGPT 進行文案生成與數據分析",
          ["文案生成實作：社群媒體文案、活動企劃案、廣告文案與影音廣告腳本",
           "數據分析實作：數據圖像化、Excel 檔整理、分析與運算，以及生產資料分析的實務建議"]),
    306: ("ChatGPT 實作", "運用 ChatGPT 進行市場調查分析", []),
    307: ("ChatGPT 實作", "運用 ChatGPT 進行數據分析與策略擬定", []),
    401: ("第六部份", "Gemini 搜索、檢索與生成技術應用",
          ["區分 AI 搜索、檢索與生成的核心差異",
           "深入介紹 Google AI 搜索與檢索、Deep Research 與 Gemini Notebook 三者構成的工作流程"]),
    407: ("Gemini 生圖", "Gemini 生圖功能與核心概念",
          ["從生圖、改圖、資訊圖表到一致性的應用指南",
           "涵蓋模型選擇、提示詞撰寫、圖片修改、風格一致性、與資訊圖表生成"]),
    412: ("Gemini Canvas", "Gemini Canvas 與創作功能",
          ["Canvas 三大功能：寫文章、生成投影片、寫程式",
           "從簡報撰作、Google Slides 生成、Viber Coding 開發，到 Apps Script 操作應用"]),
    419: ("Gemini Notebook", "認識 Gemini Notebook",
          ["9 大功能、介面三區",
           "從資料來源區、對話區到工作室：資訊圖表、簡報、資料表格、影片與語音摘要、心智圖、報告、學習卡、測驗"]),
    426: ("Gemini Notebook 職場實戰", "Gemini Notebook 職場實戰四例",
          ["競品分析簡報、報價單比較表、會議紀錄、財報解讀",
           "四個可直接套用於金屬扣件企業的工作情境：從蒐集、結構化到視覺化，並以資料表、簡報與提示詞完成交付"]),
    436: ("Gemini × Gemini Notebook", "Gemini × Gemini Notebook 市場研究工作流",
          ["8 步驟：Deep Research 蒐集完整資料，Gemini Notebook 提煉核心洞察",
           "《如何在 30 分鐘內完成 10 小時的研究》：AI 的價值不是取代創意，而是消滅「沒人真正想做」的繁瑣雜活"]),
}
BG.update({k: NAVY for k in DIVIDERS})

# 一頁拆成幾頁（原檔內容超出版面）
DUP = {304: 2,  # 廣告文案生成：步驟一、步驟二各一頁
       409: 2, 439: 2, 518: 2, 535: 2, 538: 2, 606: 2, 607: 2}  # 第二頁是 interactive.html 裡的互動演練

# 插入簡報的內容頁沒有頁首小標：依單元補上
EYEBROWS = [
    (302, 305, "第三部份｜ChatGPT 文案生成實作"), (308, 312, "第三部份｜ChatGPT 數據分析實作"),
    (402, 406, "第六部份｜Gemini 搜索與 Deep Research"), (408, 411, "第六部份｜Gemini 生圖"),
    (413, 418, "第六部份｜Gemini Canvas"), (420, 425, "第六部份｜Gemini Notebook"),
    (427, 435, "第六部份｜Gemini Notebook 職場實戰"), (437, 440, "第六部份｜Gemini × Gemini Notebook 市場研究"),
    (602, 607, "第七部份｜單元六｜Claude Skills"),
]

# 個別頁的放大上限（1.35 是預設；字被嫌太大的頁壓低）
TMAX = {89: 0.7, 96: 0.65, 101: 0.5, 422: 0, 429: 0, 433: 0, 435: 0.3, 508: 0.55, 439.1: 0.6, 606.1: 0, 605: 0.3}

CREAM, AMBER = "#F5F2EC", "#F0A35E"

# 灰字改色（看不清楚的頁）
RECOLOR = {
    20: {GREY: NAVY, DARK: NAVY},
    21: {GREY: NAVY, DARK: NAVY},
    22: {GREY: NAVY},
    24: {GREY: NAVY},
    28: {GREY: "#111111"},
    41: {GREY: NAVY},
    67: {GREY: NAVY},
}


def apply(i, els, h):
    """i：原 PPT 頁碼；els：該頁元素；h：build 模組（工具函式）"""
    one = lambda pat: h.find(els, pat)[0]
    many = lambda pat: h.find(els, pat)
    tables = [e for e in els if e["tag"] == "graphicFrame"]
    role = lambda name: next(e for e in els if e["role"] == name)

    def label(pat, x=1.6):
        """文字框第一行的小標（例如「Work 交辦指令範例」）加大"""
        one(pat)["paras"][0]["x"] = x

    if i in (20, 21):  # 來源列維持灰色
        for r in (r for p in role("source")["paras"] for r in p["runs"]):
            r["color"] = DARK

    if i in DIVIDERS:
        divider(els, *DIVIDERS[i])
    elif 300 <= i < 700:  # 插入的四份簡報的內容頁
        if 300 <= i < 500 or i >= 600:
            restyle(els, BLUE_MAP)
            # 整頁的背景圖／底色、頁首的整幅裝飾橫幅圖
            els[:] = [e for e in els if "paras" in e or not (e["w"] >= 1200 and (e["h"] >= 680 or (e["tag"] == "pic" and e["y"] <= 0)))]
        t = auto_roles(els, h)
        if i in (304.1, 304.2):  # 小標的框太窄，字會折行
            for e in many(r"^(先了解廣告文案|依據關鍵要素)"):
                e["w"] = 700
        if i == 304.1:  # 廣告文案生成（步驟一）：只留上半
            els[:] = [e for e in els if e["y"] < 455 or e["role"]]
        elif i == 304.2:  # 步驟二：拿掉上半，步驟二的標題升為頁標題
            els[:] = [e for e in els if e["role"] == "eyebrow" or e["y"] >= 455]
            t = one(r"步驟二")
            t["role"] = "title"
            t["h"] = 51
            for r in t["paras"][0]["runs"]:
                r.update(z=40.0)
            dy = t["y"] - 129
            for e in els:
                if e["role"] != "eyebrow":
                    e["y"] -= dy
        if 300 <= i < 500 or i >= 600:
            for r in t["paras"][0]["runs"]:  # 標題改成主簡報的深藍襯線體
                r.update(color=NAVY, serif=True, z=40.0, b=True)
            t["h"] = 51
        ey = next((e for e in els if e["role"] == "eyebrow"), None)
        if ey is None:
            text = next((txt for lo, hi, txt in EYEBROWS if lo <= i <= hi), "")
            ey = mk_text(t, 85.3, 85.3, 1142.7, 25.1, [text], 16.0, "#A3500C", lh=22.4)
            ey["role"] = "eyebrow"
            els.insert(0, ey)
        elif 500 <= i < 600:
            unit = {512: "單元一｜Claude AI 核心功能", 523: "單元三｜模型與方案", 524: "單元三｜模型與方案", 525: "單元三｜模型與方案",
                    526: "單元三｜模型與方案", 527: "單元三｜模型與方案", 537: "單元七", 538: "單元七｜使用注意事項", 539: "單元七｜使用注意事項"}
            if i in unit:  # 原本是描述句的頁首小標，改成單元名稱，跟其他頁一致
                h.set_text(ey["paras"][0], unit[i])
            for para in ey["paras"]:
                for r in para["runs"]:
                    if r["text"].strip():
                        r["text"] = "第七部份｜" + r["text"]
                        break
    if i == 310:  # 頁首的裝飾橫幅圖已拿掉，內容往上靠
        body = [e for e in els if not e["role"]]
        dy = min(e["y"] for e in body) - 200
        for e in body:
            e["y"] -= dy
    if i == 312:  # 原檔的四個圖示是外部連結（檔案裡沒有圖）：圓圈改成橘底編號
        circles = sorted((e for e in els if "paras" not in e and abs(e["w"] - e["h"]) < 2 and e["w"] > 40), key=lambda e: (e["y"] // 150, e["x"]))
        like = one(r"^資料準備")
        for n, c in enumerate(circles, 1):
            c["fill"], c["stroke"] = "#C2610F", None
            num = mk_text(like, c["x"], c["y"] + c["h"] / 2 - 15, c["w"], 30, [str(n)], 22.0, "#FFFFFF")
            num["paras"][0]["algn"] = "center"
            els.append(num)
    if i == 428:  # 實戰二：原檔內容到頁底，往上壓縮，底部備註字級改成跟卡片一樣
        f = 0.82
        for e in els:
            if e["role"]:
                continue
            e["y"] = 190 + (e["y"] - 190) * f
            if "paras" not in e and e["h"] > 100:
                e["h"] *= f
        note = one(r"^資安評估準則")
        note["h"] = 72
        for para in note["paras"]:
            for r in para["runs"]:
                r["z"] = 20.0
        for pat, z in ((r"^於右側「工作室」", 18.5), (r"^Gemini Notebook 沒有外部知識", 20.0)):
            for r in one(pat)["paras"][0]["runs"]:
                r["z"] = z
    if i == 429:  # 實戰三：第 2 張卡的字稍微縮小才放得進卡片
        for r in one(r"^登入 Gemini Notebook")["paras"][0]["runs"]:
            r["z"] *= 0.9
    if i == 501:  # Claude 簡報的第一頁：套用主簡報封面的版面（10/19 另一天授課）
        like = one(r"Claude AI 核心功能導覽")
        bar = {"tag": "sp", "role": "", "x": 85.3, "y": 247.7, "w": 106.7, "h": 5.3, "ph": None, "prst": "roundRect", "adj": 50000,
               "fill": AMBER, "stroke": None, "sw": 1.0, "arrow": False}
        line = {"tag": "sp", "role": "", "x": 85.3, "y": 439, "w": 1030, "h": 0, "ph": None, "prst": "line", "adj": 16667,
                "fill": None, "stroke": "#3A4A63", "sw": 1.33, "arrow": False}
        ey = mk_text(like, 85.3, 279.7, 1142.7, 34.4, ["金屬扣件AI智慧品保視覺檢測與決策應用班"], 22.7, AMBER, lh=31.7)
        ey["paras"][0]["runs"][0]["spc"] = 2.67
        t = mk_text(like, 85.3, 338, 1142.7, 76.3, ["Claude 功能與使用全攻略"], 64, CREAM, serif=True, lh=73.6)
        who = mk_text(like, 85.3, 457.7, 1030, 104, ["講者：國立澎湖科技大學 行銷與物流管理系 唐嘉偉", "授課日期：115/10/19", "授課地點：大台南會展中心"],
                      18.7, CREAM, b=False, lh=32)
        for para in who["paras"][1:]:
            para["sb"] = 5.3
            para["runs"][0]["color"] = AMBER
        who["boost"] = 1.7
        els[:] = [bar, ey, t, line, who]

    if isinstance(i, float) and int(i) != 304 and round(i % 1, 1) == 0.1:
        i = int(i)  # DUP 拆出的第一頁（409.1 等）沿用整數頁碼的調整

    if i == 413:  # Canvas 功能總覽：三張卡片的字距、行距拉開
        for e in many(r"^(協助撰寫|依大綱|以自然語言)"):
            e["y"] += 22
            for para in e["paras"]:
                para["lh"] = round(para["lh"] * 1.3, 1)
    if i == 418:  # Apps Script 部署六步驟：(A)(B)(C) → (1)(2)(3)，字縮小並對齊
        for e in els:
            if "paras" not in e or e["role"]:
                continue
            for para in e["paras"]:
                for r in para["runs"]:
                    r["text"] = re.sub(r"\(([ABC])\)", lambda m: "(%d)" % ("ABC".index(m.group(1)) + 1), r["text"])
                    if r["z"] > 23 and r["color"] != "#A3500C":  # 內文
                        r["z"] = 20.0
                    elif r["z"] > 23:  # 卡片標題
                        r["z"] = 22.0
            if any(re.match(r"\(1\)", para["runs"][0]["text"]) for para in e["paras"]):
                e["w"], e["y"] = 322, e["y"] + 6
                for k, para in enumerate(e["paras"]):
                    para["hang"], para["sb"], para["lh"] = 1.9, (6 if k else 0), 29
    if i == 421:  # Gemini Notebook 介面三區：三張卡片改成條列
        bullets(one(r"^支援 PDF"), ["支援 PDF、純文字、YouTube 影片、網址、Google 雲端硬碟", "顯示文件列表，可隨時新增內容",
                                   "整理後於問答區呈現小總結，可儲存成「記事」"])
        bullets(one(r"^針對想深入的問題"), ["針對想深入的問題提問，AI 也會提供參考問題", "回覆以列點呈現並附上來源，方便核對原始資料",
                                        "對話紀錄自動儲存、可刪除；共享筆記本中的聊天紀錄僅自己可見",
                                        "「自訂」聊天：設定目標、語氣或角色，例如「Treat me like a PhD candidate」"])
        bullets(one(r"^音訊（Podcast）"), ["音訊（Podcast）摘要、影片摘要", "心智圖、報告、學習卡、測驗", "資訊圖表、簡報、資料表格",
                                        "依自己消化資訊的方式選擇適合的功能"])
    if i == 422:  # 實戰教學：介面介紹：截圖縮小一點，右側卡片與底部「扣件企業對應」往上
        pic = next(e for e in els if e["tag"] == "pic" and e["w"] > 500)
        pic["w"], pic["h"] = round(783 * 364 / 399), 364
        tops = {"① 新增來源": (256, 104), "② 對話提問": (372, 104), "③ 工作室產出": (488, 136)}
        for head, (top, ch) in tops.items():
            hd = one("^" + head)
            dy = top + 18 - hd["y"]
            body = min((e for e in els if "paras" in e and e is not hd and e["x"] > 900 and e["y"] > hd["y"]), key=lambda e: e["y"])
            card = min((e for e in els if "paras" not in e and e["x"] > 900 and e["h"] > 50 and abs(e["y"] + 18 - hd["y"]) < 30), key=lambda e: e["y"])
            bar = next(e for e in els if "paras" not in e and e["x"] > 900 and e["h"] < 12 and abs(e["y"] - card["y"]) < 6)
            card["y"], card["h"], bar["y"] = top, ch, top - 2
            hd["y"], body["y"], body["h"] = top + 16, top + 46, ch - 52
            for r in runs_of(body):
                r["z"] = 19.0
            for para in body["paras"]:
                para["lh"] = 27
        box = one(r"^扣件企業對應")
        bg = next(e for e in els if "paras" not in e and e.get("fill") == "#EFE3D3" and e["w"] > 1000)
        icon = next((e for e in els if e["tag"] == "pic" and e["w"] < 40), None)
        bg["y"], bg["h"] = 640, 70
        box["y"], box["h"] = 648, 58
        for para in box["paras"]:
            para["lh"] = 30
        for r in runs_of(box):
            r["z"] = 21.0
        if icon:
            icon["y"] = 664
    if i == 429:  # 實戰三：左側提示詞字級統一、右側三張卡片收緊並上移、底部提醒上移
        prompt = one(r"^會議紀錄提示詞")
        prompt["w"] = 662
        for k, para in enumerate(prompt["paras"]):
            para["lh"], para["sb"] = 29, (4 if k == 1 else 2 if k else 0)
            para["hang"] = 1.5 if re.match(r"\d\.", para["runs"][0]["text"]) else None
            for r in para["runs"]:
                r["z"] = 20.0
                if k == 0:
                    r["z"], r["color"] = 21.0, "#A3500C"
                elif r["color"] == "#15233A":
                    r["b"] = False
        for cy0, ctop, ch in ((257, 257, 84), (389, 363, 136), (561, 521, 132)):  # 原卡片上緣 → 新上緣、新高度
            group = [e for e in els if e["x"] >= 795 and cy0 - 24 <= e["y"] <= cy0 + 62]
            for e in group:
                e["y"] += ctop - cy0
            card = next(e for e in group if "paras" not in e and e["h"] > 50)
            card["h"] = ch
            hd = next(e for e in group if "paras" in e and e["h"] < 30 and e["y"] < ctop + 40)
            body = next(e for e in group if "paras" in e and e is not hd and not re.match(r"^\d$", h.plain(e)))
            body["x"], body["w"], body["y"], body["h"] = 815, 368, ctop + 54, ch - 60
            for para in body["paras"]:
                para["lh"] = 24
            for r in runs_of(body):
                r["z"] = 17.0
        note = one(r"^人工覆核")
        bg = next(e for e in els if "paras" not in e and e.get("fill") == "#EFE3D3" and e["w"] > 1000 and e["y"] > 600)
        icon = next((e for e in els if e["tag"] == "pic" and e["y"] > 600), None)
        bg["y"], bg["h"], note["y"] = 672, 54, 687
        if icon:
            icon["y"] = 689
        for r in runs_of(note):
            r["z"] = 23.0
    if i == 430:  # 實戰四：四張卡片改成條列
        bullets(one(r"^AI 不會直接推薦股票"), ["AI 不會直接推薦股票，而是幫你「搞懂財報到底在說什麼」，把長篇財報整理成好消化的重點摘要",
                                          "推薦 Gemini Notebook 的原因：只根據上傳文件回答、不無中生有；可長期追蹤單一公司的財報變化"])
        bullets(one(r"^台股："), ["台股：公開資訊觀測站下載「財務報告書」與「法說會簡報」", "美股：Google 財經搜尋個股，於「收益」分頁下載相關文件"])
        bullets(one(r"^於筆記本左側"), ["於筆記本左側來源區上傳檔案", "建議同時上傳前一季財報，以便比較"])
        bullets(one(r"^請 AI 扮演資深分析師"), ["請 AI 扮演資深分析師，針對營收、毛利率等關鍵指標分析", "系統產出結構清晰的財報分析報告（提示詞見下頁）"])
    if i == 431:  # 財報分析提示詞：內文不要粗體
        for e in many(r"^(入門版|進階版)"):
            for para in e["paras"][1:]:
                for r in para["runs"]:
                    r["b"] = False
    if i == 433:  # 先建可信資料庫：六張卡片的標題一行、內文從同一高度開始
        for e in els:
            if "paras" not in e or e["role"]:
                continue
            txt = h.plain(e)
            if re.match(r"^(0\d |扣件企業對應)", txt):
                e["w"], e["h"] = 322, 58
                for r in runs_of(e):
                    r["z"] = 19.5
                e["paras"][0]["lh"] = 27
            elif re.match(r"^(用途|公司簡介)", txt):
                for k, para in enumerate(e["paras"]):
                    para["sb"] = 6 if k else 0
                for r in runs_of(e):
                    r["z"] = 19.5
                for para in e["paras"]:
                    para["lh"] = 28
        for card_top, body_top in ((269, 347), (505, 583)):
            for e in els:
                if "paras" in e and not e["role"] and re.match(r"^(用途|公司簡介)", h.plain(e)) and card_top < e["y"] < card_top + 200:
                    e["y"] = body_top
    if i == 435:  # 從資料到提案的 5 步驟：三欄對齊、Step 標題改深藍
        for e in els:
            if "paras" not in e or e["role"]:
                continue
            txt = h.plain(e)
            row = int((e["y"] - 213) // 104)
            top = 213 + row * 104
            if txt.startswith("Step"):
                e.update(x=185, w=330, y=top + 18, h=54, anchor="ctr")
                for r in runs_of(e):
                    r["z"], r["color"] = 21.0, NAVY
            elif txt.startswith("操作"):
                e.update(x=520, w=372, y=top + 18, h=54, anchor="ctr")
                for r in runs_of(e):
                    r["z"] = 19.0
            elif txt.startswith("產出"):
                e.update(x=898, w=290, y=top + 18, h=54, anchor="ctr")
                for r in runs_of(e):
                    r["z"] = 19.0
    if i == 439:  # 8 步驟提示詞範例：兩個框收緊，底部的扣件產業案例回到版面內
        boxes = sorted((e for e in els if "paras" not in e and e.get("fill") == "#EFE3D3"), key=lambda e: e["y"])
        icons = sorted((e for e in els if e["tag"] == "pic"), key=lambda e: e["y"])
        t1, t7, case = one(r"^第 1 步"), one(r"^第 7 步"), one(r"^扣件產業案例")
        boxes[0]["y"], boxes[0]["h"] = 240, 214
        t1["y"], t1["h"] = 248, 200
        boxes[1]["y"], boxes[1]["h"] = 466, 150
        t7["y"], t7["h"] = 474, 136
        case["y"], case["h"] = 628, 56
        for e in (t1, t7, case):
            for para in e["paras"]:
                para["lh"] = 30
            for r in runs_of(e):
                r["z"] = 20.0
        for ic, y in zip(icons, (256, 482)):
            ic["y"] = y
    if i == 602:  # 第二個標題改成副標（橘色、26px）
        sub = one(r"^一個 Claude Skill 的聘書")
        sub["h"] = 34
        for r in runs_of(sub):
            r.update(z=26.0, color="#A3500C", serif=False)
    if i == 606:  # 詢價單整理小幫手：兩欄文字都偏多，縮小字級才放得進卡片；卡片底部收回版面內
        for e in els:
            if "paras" in e and not e["role"] and e["h"] > 100:
                for para in e["paras"]:
                    para["lh"] = 27
                for r in runs_of(e):
                    r["z"] = 18.5
            if not e["role"] and e["h"] > 250:
                e["h"] = min(e["h"], 700 - e["y"])
    if i == 508:  # 擴充功能：三張卡片的「應用情境」對齊同一高度
        for e in els:
            if e.get("fill") == "#EFE3D3" and e["w"] < 400:
                e["y"] = 448
            elif "paras" in e and h.plain(e).startswith("應用情境"):
                e["y"] = 456
    if i == 527:  # Claude 可以生成圖片嗎：兩張卡片的內文放大
        for e in many(r"^(與其他 AI 生圖工具|分析使用者上傳)"):
            for para in e["paras"]:
                para["lh"] = 38
            for r in runs_of(e):
                r["z"] = 24.0

    if i == 201:  # 插入頁：常見的生成式 AI 三大功能（1600×900 來源，已換算成 1280 寬）
        t = one("常見的 生成式 AI 三大功能")
        t["role"] = "title"
        for e in els:
            e["y"] += 75  # 這份簡報的標題貼近上緣，先往下挪到與主簡報相同的位置（標題 y≈129）
        for r in t["paras"][0]["runs"]:  # 標題改成主簡報的深藍襯線體
            r["color"], r["serif"], r["z"] = NAVY, True, 40.0
        ey = copy.deepcopy(t)  # 補上鄰頁都有的頁首小標
        ey.update(role="eyebrow", y=85.3, h=25.1, x=85.3, w=1142.7)
        h.set_text(ey["paras"][0], "第一部份｜AI工具指南")
        ey["paras"][0]["lh"] = 22.4
        for r in ey["paras"][0]["runs"]:
            r.update(color="#A3500C", serif=False, z=16.0, b=True, spc=1.33)
        els.insert(0, ey)
        # 第三張卡的文字框在原檔只有半寬，字會折成很多行：改成跟前兩張一樣寬
        for pat, ref in (("功能：AI 生成", "功能：AI 搜索"), ("依資料重組", "從全球網路")):
            one(pat)["w"] = one(ref)["w"]

    if 60 <= i <= 79:  # 第二部份拿掉後，原第三部份改為第二部份
        for e in els:
            for para in e.get("paras", []):
                for r in para["runs"]:
                    r["text"] = r["text"].replace("第三部份", "第二部份")
    if i == 95:  # 先講 Skills 再講 Work
        h.set_text(one(r"ChatGPT：Work 與 Skills")["paras"][0], "ChatGPT：Skills 與 Work")
        one(r"^ChatGPT Work：")["paras"].reverse()

    if i == 1:
        h.set_text(one(r"^金屬扣件智慧製造")["paras"][0], "金屬扣件AI智慧品保視覺檢測與決策應用班")
        h.set_text(one(r"^AI倫理、資安與法規介紹")["paras"][0], "生成式AI發展與金屬扣件產業轉型")
        h.remove(els, one(r"^生成式AI應用流程與AI Agent"))
        for e in els:  # 副標拿掉後，分隔線與講者資訊往上靠
            if e["y"] >= 495:
                e["y"] -= 60
        e = one(r"^講者：")
        h.set_text(e["paras"][1], "授課日期：115/10/15")
        h.set_text(e["paras"][2], "授課地點：大台南會展中心")
        e["boost"] = 1.7
        for para in e["paras"][1:]:  # 授課日期、授課地點換成琥珀色，和講者那一行區分
            for r in para["runs"]:
                r["color"] = AMBER
    elif i == 10:
        h.recolor_pills(els, NAVY, CREAM)
    elif i == 7:
        h.add_para(one(r"^Strategy Dynamics"),
                   "Generative AI Foundations（2026/07/22；發照機構：Microsoft Certified Fundamentals）")
    elif i == 15:
        tables[0]["boost"] = 1.2
    elif i == 17:
        one(r"^Mollick指出")["boost"] = 1.3
    elif i == 19:
        for e in many(r"^許多使用者|^對追求效益"):
            e["boost"] = 1.2
        for e in many(r"^\(\d\) "):
            h.cut(e, r"^\(\d\) ")
        h.recolor_pills(els, NAVY, CREAM)
    elif i == 20:
        tables[0]["boost"] = 1.15
    elif i == 21:
        lead = one(r"^ChatGPT、Gemini、Claude，三大AI工具怎麼選？（續）")
        h.remove(els, lead)
        top = min(e["y"] for e in els if not e["role"])
        for e in els:  # 下方卡片往上遞補
            if not e["role"]:
                e["y"] -= top - lead["y"]
    elif i == 22:
        for e in many(r"^無論免費版|^當時有Gemini|^被譽為"):
            e["boost"] = 1.15
        e = one(r"^無論免費版")
        h.mark(e, "並非單一模型", h.RED)
        h.mark(e, "手動選擇GPT-5.2 Thinking Extended")
        h.mark(e, "可選擇GPT-5.2 Pro")
        e = one(r"^當時有Gemini")
        h.mark(e, "三種版本可供選擇")
        h.mark(e, "Gemini Deep Think")
        h.mark(e, "建議選擇Gemini 3 Pro或Thinking版本")
        e = one(r"^被譽為")
        h.mark(e, "最有「人味」的AI")
        h.mark(e, "Opus 4.6版本，並啟用「擴展思維」功能")
        e = one(r"^共同原則")
        h.mark(e, "簡單任務用預設模型即可", h.GREEN)
        h.mark(e, "手動切換至「思考／推理」版本")
        h.mark(e, "給予AI充分的背景資訊")
    elif i == 23:
        h.cut(one(r"^在代理時代"), r"本課程對應：.*?路徑。")
        e = one(r"^\(1\) 若尚未")  # 「導入步驟」改用和左右兩張卡片一樣的圓點
        h.cut(e, r"^\(\d\) ")
        for para in e["paras"]:
            para["bu"] = one(r"^不要只是把AI")["paras"][0]["bu"]
    elif i == 28:
        h.sub(role("source"), "NIST AI RMF 1.0（AI 100-1）", "美國國家標準暨技術研究院(NIST)人工智慧風險管理框架1.0（AI 100-1）")
    elif i == 30:
        # 後來貼上的超連結文字框疊在清單上：併進清單（原本留了一個空的項目符號給它）
        link, host = one(r"^https://nvlpubs"), one(r"^2023-01-26")
        blank = next((p for p in host["paras"] if not h.ptext(p).strip()), None)
        if blank is None:
            blank = copy.deepcopy(host["paras"][-1])
            host["paras"].append(blank)
        blank["runs"] = copy.deepcopy(link["paras"][0]["runs"])
        host["free"] = True
        h.remove(els, link)
    elif i == 29:
        h.rm_card(els, r"^2027-12-02")
        h.respread(els, [c for c in els if "paras" not in c and c.get("fill") and abs(c["y"] - 198) < 3 and c["h"] > 100])
        for e in many(r"^正式生效|^禁止行為|^通用目的|^嵌入機械"):
            e["boost"] = 1.15
        h.stack(els, many(r"^不可接受風險|^高風險：|^透明度風險|^最低風險"))
        h.stack(els, many(r"^第 2 條涵蓋|^含 AI 功能|^單純廠內"))
    elif i == 33:
        h.remove(els, one(r"^2020 年增訂"))
    elif i == 34:
        h.remove(els, one(r"^子法草案"), one(r"^外洩後 72 小時"),
                 *[e for e in els if e.get("prst") == "line" and abs(e["y"] - 490) < 4 and e["h"] > 20])
    elif i == 35:
        one(r"^AI 時代的合理保密措施")["boost"] = 1.3
        green, red = one(r"^可使用核准工具"), one(r"^禁止輸入外部 AI")
        green["y"], green["h"] = red["y"], red["h"]  # 綠框和紅、黃框一樣大
    elif i == 39:
        # 第三張卡片的兩個文字框原本超出卡片，併成一個並收回卡片內
        a, b = one(r"^Stephen Thaler"), one(r"^上訴法院")
        para = copy.deepcopy(b["paras"][0])
        para["sb"] = 10
        a["paras"].append(para)
        a["x"], a["w"], a["h"] = 859, 323, 222
        a["free"] = True  # 高度是估的，交給排版引擎依實際字型量測
        h.remove(els, b)
        h.mark(one(r"^經授智字"), "使用者下指令的AI內容，不受保護", h.RED)
        h.mark(one(r"^將著作輸入"), "須授權或合理使用")
        h.mark(one(r"^報告第 2 部"), "不足以取得作者身分與著作權保護", h.RED)
        h.mark(one(r"^現行法律原則"), "再進行創意編排或修改者，則可受保護", h.GREEN)
        h.mark(a, "將 AI 列為作者")
        h.mark(a, "著作權須以人類為作者", h.RED)
        h.mark(a, "判決確定")
    elif i == 40:
        tables[0]["boost"] = 1.15
    elif i == 41:
        marks = [(r"^輸入內容是否有權使用", ["是否有權使用"]),
                 (r"^輸出是否疑似", ["重現既有作品、商標或浮水印"]),
                 (r"^引用的數據", ["回查原始出處"]),
                 (r"^對外文件是否", ["人工修改", "標示 AI 協作"]),
                 (r"^使用的工具方案", ["智財補償"]),
                 (r"^需主張權利的內容", ["保留人類創作紀錄"])]
        for pat, phrases in marks:
            e = one(pat)
            e["w"] = 618  # 原本依字數各自縮短，統一寬度才不會放大後折行
            for ph in phrases:
                h.mark(e, ph)
    elif i == 44:
        tables[0]["boost"] = 1.15
    elif i == 45:
        for e in many(r"^後果："):
            h.cut(e, r"^後果：")
            e["boost"] = 1.3
    elif i == 49:
        t = tables[0]
        ratio = 1107 / sum(t["cols"])  # 表格原本偏窄，拉到與版心同寬
        t["cols"] = [c * ratio for c in t["cols"]]
        t["boost"] = 1.15
    elif i == 62:
        for e in many(r"^Prompt（提示詞）|^OpenAI：|^回答與摘要"):
            e["boost"] = 1.15
    elif i == 66:
        e = one(r"^為什麼要結構化")
        h.set_text(e["paras"][1], "只問「請幫我整理這份資料」，AI的回答不一定穩定、可檢查，也接不上下一個流程。")
        h.set_text(e["paras"][2], "企業需要的AI產出：可預期、可追溯、可驗收、可延伸。")
        e["paras"][3]["x"] = 1.3
        e["boost"] = 2.0
        h.shrink_row(els, [e], 136)
        for e in many(r"^目的："):
            e["boost"] = 1.3
    elif i == 67:
        e = one(r"^你是一位企業內部")
        h.mark(e, "企業內部知識管理助理")
        h.mark(e, "不是自由聊天", h.RED)
        h.mark(e, "可執行、可檢查、可交接", h.GREEN)
        e = one(r"^請根據我提供的公司文件")
        h.mark(e, "標準答案、依據來源、操作步驟與注意事項")
        h.mark(e, "新人可以照著執行", h.GREEN)
        e = one(r"^只能根據提供的文件")
        h.mark(e, "只能根據提供的文件、表格或知識庫回答")
        h.mark(e, "不要自行推測或編造", h.RED)
        h.mark(e, "向主管或相關窗口確認", h.GREEN)
    elif i == 68:
        e = one(r"^結構化Prompt的價值，是")
        h.set_text(e["paras"][0], "把靠人腦臨場判斷、品質不穩定的任務，轉成LLM能穩定承接、重複產出的工作規格；這是個人使用與企業應用的分水嶺。")
        h.set_text(e["paras"][1], "工具會換、模型會換，任務不會消失：整理資料、回覆客戶、產出報告、檢查錯誤、輔助決策。")
        e["boost"] = 1.2
        h.shrink_row(els, [e], 66)
    elif i == 72:
        for e in many(r"可直接套用的內部助理指令|標準回答格式|驗收題目"):
            e["boost"] = 1.1
    elif i == 75:
        one(r"^優化完成後")["boost"] = 1.3
    elif i == 76:
        label(r"^優化後的完整版Prompt")
    elif i == 77:
        label(r"^個人化建議指令")
        one(r"^透過三項個人化建議")["boost"] = 1.6
    elif i in (78, 90, 91, 93, 94):
        h.merge_heads(els)
    elif i == 83:
        h.merge_heads(els, keep_num=False)  # 拿掉「1. 明確問題意識」上方多餘的 01、02
    elif i == 85:
        label(r"^Deep Research提問範例")
    elif i == 89:
        e = one(r"^隨著ChatGPT Image 2.0")
        h.set_text(e["paras"][0], "ChatGPT Image 2.0上線後，圖像的質感、細節與指令理解明顯提升；只要給對提示詞，就能更快產出接近需求的圖片。")
        e["boost"] = 1.2
        e = one(r"^提示詞的共同結構")
        h.sub(e, h.ptext(e["paras"][0])[len("提示詞的共同結構："):],
              "用途＋主題＋風格與色調＋版面要求＋使用情境。扣件企業可套用於產品型錄、展會海報、客戶簡報與流程圖解。")
        e["boost"] = 1.2
        h.shrink_row(els, [one(r"^ChatGPT Image 2.0上線後"), e], 48)
    elif i == 92:
        label(r"^長提示詞範例")
        for e in many(r"^長提示詞的四個層次|^扣件產業改寫"):
            e["boost"] = 2.5
    elif i == 96:
        tables[0]["boost"] = 1.15
    elif i == 97:
        lists = {
            r"^直接產出試算表": ["直接產出試算表、文件與簡報。", "可產出用網址分享的互動網站 Sites（儀表板、追蹤表、報告）。", "資料更新後可同步更新內容。"],
            r"^透過外掛程式": ["透過外掛程式（plugins）連接工作應用程式。", "可連接 Slack、Microsoft Teams、Google Drive、SharePoint、電子郵件、行事曆、CRM 與專案追蹤工具。", "在指令中以「@」指定工具。"],
            r"^把複雜任務拆成": ["把複雜任務拆成多個步驟依序完成。", "過程中保留前後脈絡。", "使用者可隨時檢視進度並調整方向。"],
            r"^可設定單次": ["可設定單次或定期執行。", "連結的應用程式發生特定事件時自動執行。", "結果集中在「Scheduled」分頁檢視。"],
            r"^桌面版可在": ["桌面版可在使用者允許下讀取本機檔案與應用程式。", "透過內建瀏覽器與 Computer Use 點選、輸入與搬移檔案。"],
        }
        for pat, items in lists.items():
            h.set_list(one(pat), items)
        h.merge_heads(els, sep=" ")  # 「01／產出完成的檔案」併成一行，把高度留給條列
    elif i == 100:
        h.cut(role("source"), r"原簡報")
    elif i == 101:
        cell = next(c for c in tables[0]["rows"][0]["cells"] if "原簡報要點" in "".join(h.ptext(p) for p in c["paras"]))
        h.sub(cell, "原簡報要點", "要點")
    elif i == 102:
        label(r"^Work 交辦指令範例")
    elif i == 103:
        for e in many(r"^提示詞 \d"):
            h.cut(e, r"^提示詞 \d、\d：")
            e["boost"] = 1.15
    elif i == 104:
        h.sub(one(r"^扣件廠的四個"), "四個", "三個")
        h.sub(one(r"^把案例方法套用到"), "業務、品管、生管與經營管理", "業務、品管與經營管理")
    elif i == 105:
        # 「流程三：生產週報」那一頁已刪，總覽改成三個流程
        h.sub(role("title"), "四個工作流程總覽", "三個工作流程總覽")
        t = tables[0]
        text = lambda row: "".join(h.ptext(p) for c in row["cells"] for p in c["paras"])
        t["rows"] = [r for r in t["rows"] if "生產週報與稼動分析" not in text(r)]
        t["h"] = sum(r["h"] for r in t["rows"])
        cell = next(c for r in t["rows"] for c in r["cells"] if "前三項成果" in "".join(h.ptext(p) for p in c["paras"]))
        h.sub(cell, "前三項成果", "前兩項成果")
    elif i == 109:
        h.sub(role("title"), "流程四", "流程三")
        h.sub(one(r"^對應原簡報的"), "對應原簡報的", "對應")
        h.sub(one(r"^匯入前三個流程"), "匯入前三個流程的詢價彙整表、品質分析結果與生產週報。", "匯入前兩個流程的詢價彙整表與品質分析結果。")
        h.sub(one(r"^Work 交辦指令範例"), "請以本月的生產週報、品質分析結果與詢價彙整表為依據", "請以本月的品質分析結果與詢價彙整表為依據")
    elif i == 112:
        h.cut(role("eyebrow"), r"：原簡報內容")
        h.cut(role("source"), r"原簡報")
    elif i == 116:
        h.rm_card(els, r"^官方定義")
    elif i == 117:
        h.rm_card(els, r"^官方原則")
    elif i == 118:
        h.merge_heads(els, sep=" ")
    elif i in (120, 121, 122, 123, 124):
        h.cut(role("eyebrow"), r"（原簡報.*?改寫為 Skill）")
        h.cut(one(r"^原 GPTs 指令"), r"（原簡報）")
        h.rm_card(els, r"^改寫效益|^扣件廠應用|^改寫重點")
        h.cut(role("source"), r"案例來源：原簡報〈.*?〉；")
    elif i == 127:
        h.rm_card(els, r"^資料使用")
