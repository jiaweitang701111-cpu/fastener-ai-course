# -*- coding: utf-8 -*-
"""依「原 PPT 頁碼」做的內容調整。build.py 讀完每一頁後會呼叫 apply()。

PPT 改版後若頁序變動，這裡的頁碼要跟著改；找不到指定文字時 build 會停下來並指出是哪一頁。
"""
import copy
import re

GREY, DARK, NAVY = "#3E4756", "#5F6673", "#15233A"

# 整頁刪除
DROP = {18, 69, 73, 108, 111, 115}

# 個別頁的放大上限（1.35 是預設；字被嫌太大的頁壓低）
TMAX = {89: 0.7, 96: 0.65, 101: 0.5}

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

    if i == 1:
        e = one(r"^講者：")
        e["boost"] = 1.7
        for para in e["paras"][1:]:  # 授課日期、授課地點換成琥珀色，和講者那一行區分
            for r in para["runs"]:
                r["color"] = AMBER
    elif i == 8:
        h.sub(one(r"扣件廠四個工作流程"), "扣件廠四個工作流程", "扣件廠三個工作流程")
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
