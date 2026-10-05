# -*- coding: utf-8 -*-
"""PPTX -> 單檔網頁簡報。

用法： python build.py <來源.pptx> [輸出.html]

流程：
1. 逐頁把 PPTX 的形狀、文字、表格讀成元素（座標換算到 1280x720 舞台）。
2. patches.py 依頁碼做內容調整（刪頁、刪卡片、改字、改色、重點標註、指定加大）。
3. 內容頁做版面分析：把垂直方向切成「帶」，算出每個文字框往下可以長到哪裡。
   瀏覽器端（template.html 的 fit）據此把字放到排得下的最大，並把下方內容往下推。
4. 演練頁以 interactive.html 的互動版取代。
"""
import base64
import copy
import html
import re
import sys
import zipfile
from pathlib import Path

from lxml import etree

HERE = Path(__file__).parent
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
RID = "{%s}id" % NS["r"]
REMBED = "{%s}embed" % NS["r"]
EMU = 9525.0  # EMU per px（13.333in 寬 = 1280px）
PT = 4 / 3  # px per pt

SHIFT_UP = 50  # 內容頁整體上移（頁首貼近上緣），讓出更多內容高度
LEFT, RIGHT = 48, 1232  # 內容頁版心左右緣（原簡報是 85～1194，兩側留白較多）
WIDEN = (RIGHT - LEFT) / (1194.4 - 85.3)
BOTTOM_FREE = 678  # 沒有來源列時，內容可用到的最低位置
BODY_PX = 23.7  # 內文放大的目標字級（約 1.85u）
HEAD_PX = 26.5  # 卡片標題、引言放大的目標字級（約 2.07u）
EYEBROW_K = 1.15  # 頁首小標
SOURCE_K = 0.85  # 頁尾來源列
NAVY, RUST, RED, GREEN = "#15233A", "#C2610F", "#B3261E", "#2E7D4F"


def n(v):
    """數值轉精簡字串"""
    s = "%.1f" % v
    return s[:-2] if s.endswith(".0") else s


def gain(z):
    """字級 z(px) 最多可以再放大多少（0.48 代表 1.48 倍）"""
    g = BODY_PX / z if z < 21 else max(HEAD_PX / z, 1.1)
    return round(min(max(g, 1.0), 1.5) - 1, 3)


class Styles:
    """把重複的行內樣式收斂成 class"""

    def __init__(self, prefix):
        self.prefix, self.map = prefix, {}

    def cls(self, css):
        if css not in self.map:
            self.map[css] = "%s%d" % (self.prefix, len(self.map))
        return self.map[css]

    def css(self):
        return "".join(".%s{%s}" % (c, s) for s, c in self.map.items())


RUNS, PARAS = Styles("r"), Styles("p")


# ───────────────────────── 讀取 PPTX ─────────────────────────

def color(node, path="a:solidFill/a:srgbClr"):
    c = node.find(path, NS) if node is not None else None
    return "#" + c.get("val") if c is not None else None


def parse_run(rpr, recolor):
    r = {"z": 24.0, "b": False, "i": False, "u": False, "spc": 0, "color": None, "serif": False, "href": None}
    if rpr is not None:
        if rpr.get("sz"):
            r["z"] = int(rpr.get("sz")) / 100 * PT
        r["b"] = rpr.get("b") == "1"
        r["i"] = rpr.get("i") == "1"
        r["u"] = rpr.get("u") not in (None, "none")
        r["spc"] = int(rpr.get("spc") or 0) / 100 * PT
        c = color(rpr)
        r["color"] = recolor.get(c.upper(), c) if c else None
        latin = rpr.find("a:latin", NS)
        r["serif"] = latin is not None and "Serif" in latin.get("typeface", "")
    return r


def parse_paras(txbody, rels, recolor, scale=1):
    out = []
    for p in txbody.findall("a:p", NS):
        ppr = p.find("a:pPr", NS)
        first = p.find("a:r/a:rPr", NS)
        if first is None:
            first = p.find("a:endParaRPr", NS)
        size = int(first.get("sz")) / 100 if first is not None and first.get("sz") else 18
        para = {"algn": None, "lh": size * 1.2 * PT, "sb": 0, "bu": None, "runs": [], "x": None, "hang": None}
        if ppr is not None:
            para["algn"] = {"ctr": "center", "r": "right"}.get(ppr.get("algn"))
            sp, pc = ppr.find("a:lnSpc/a:spcPts", NS), ppr.find("a:lnSpc/a:spcPct", NS)
            if sp is not None:
                para["lh"] = int(sp.get("val")) / 100 * PT
            elif pc is not None:
                para["lh"] = size * 1.2 * PT * int(pc.get("val")) / 100000
            bef = ppr.find("a:spcBef/a:spcPts", NS)
            if bef is not None:
                para["sb"] = int(bef.get("val")) / 100 * PT
            bu = ppr.find("a:buChar", NS)
            if bu is not None:
                para["bu"] = (bu.get("char"), int(ppr.get("marL", 0)) / EMU)
        para["lh"] *= scale
        for ch in p:
            tag = etree.QName(ch).localname
            if tag == "r":
                rpr = ch.find("a:rPr", NS)
                r = parse_run(rpr, recolor)
                r["z"] *= scale
                r["text"] = ch.findtext("a:t", "", NS)
                link = rpr.find("a:hlinkClick", NS) if rpr is not None else None
                if link is not None:
                    r["href"] = rels.get(link.get(RID))
                para["runs"].append(r)
            elif tag == "br":
                r = parse_run(first, recolor)
                r["text"] = "\n"
                para["runs"].append(r)
        if not para["runs"]:
            r = parse_run(first, recolor)
            r["text"] = " "
            para["runs"].append(r)
        out.append(para)
    return out


def ptext(para):
    return "".join(r["text"] for r in para["runs"])


def plain(e):
    return "\n".join(ptext(p) for p in e.get("paras", [])).strip()


def parse_slide(xml, rels, recolor):
    """回傳 (背景色, 元素清單)。元素是 dict，座標單位 px。"""
    root = etree.fromstring(xml)
    bg = color(root.find("p:cSld/p:bg/p:bgPr", NS)) or "#F5F2EC"
    els = []
    for node in root.find("p:cSld/p:spTree", NS):
        tag = etree.QName(node).localname
        if tag not in ("sp", "cxnSp", "pic", "graphicFrame"):
            continue
        xf = node.find("p:xfrm", NS) if tag == "graphicFrame" else node.find("p:spPr/a:xfrm", NS)
        off, ext = xf.find("a:off", NS), xf.find("a:ext", NS)
        e = {"tag": tag, "role": "",
             "x": int(off.get("x")) / EMU, "y": int(off.get("y")) / EMU,
             "w": int(ext.get("cx")) / EMU, "h": int(ext.get("cy")) / EMU}
        ph = node.find(".//p:nvPr/p:ph", NS)
        e["ph"] = ph.get("type") if ph is not None else None
        sppr = node.find("p:spPr", NS)
        if tag in ("sp", "cxnSp"):
            geom = sppr.find("a:prstGeom", NS)
            e["prst"] = geom.get("prst") if geom is not None else "rect"
            gd = sppr.find("a:prstGeom/a:avLst/a:gd", NS)
            e["adj"] = int(gd.get("fmla").split()[1]) if gd is not None else 16667
            e["fill"] = color(sppr)
            ln = sppr.find("a:ln", NS)
            e["stroke"] = color(ln)
            e["sw"] = max(1.0, int(ln.get("w", 9525)) / 12700 * PT) if ln is not None else 0
            e["arrow"] = ln is not None and ln.find("a:tailEnd", NS) is not None
            tx = node.find("p:txBody", NS)
            if tx is not None and "".join(tx.itertext()).strip():
                bp = tx.find("a:bodyPr", NS)
                e["pad"] = [int(bp.get(k, d)) / EMU for k, d in
                            (("tIns", 45720), ("rIns", 91440), ("bIns", 45720), ("lIns", 91440))]
                e["anchor"] = bp.get("anchor", "t")
                na = bp.find("a:normAutofit", NS)
                scale = int(na.get("fontScale")) / 100000 if na is not None and na.get("fontScale") else 1
                e["paras"] = parse_paras(tx, rels, recolor, scale)
                if not e["fill"] and e["x"] < 1228 < e["x"] + e["w"]:
                    e["w"] = 1228 - e["x"]  # 文字框不超出右側版心
            elif tag == "sp" and e["prst"] != "line" and not e["fill"] and not e["stroke"]:
                continue  # 空的透明框
        elif tag == "pic":
            e["img"] = rels.get(node.find("p:blipFill/a:blip", NS).get(REMBED))
        else:
            tbl = node.find(".//a:tbl", NS)
            e["cols"] = [int(c.get("w")) / EMU for c in tbl.findall("a:tblGrid/a:gridCol", NS)]
            e["rows"] = []
            for tr in tbl.findall("a:tr", NS):
                cells = []
                for tc in tr.findall("a:tc", NS):
                    pr = tc.find("a:tcPr", NS)
                    css = []
                    if pr is not None:
                        css.append("padding:%spx %spx %spx %spx" % tuple(
                            n(int(pr.get(k, d)) / EMU) for k, d in
                            (("marT", 45720), ("marR", 91440), ("marB", 45720), ("marL", 91440))))
                        if pr.get("anchor") == "ctr":
                            css.append("vertical-align:middle")
                        for side, key in (("left", "lnL"), ("right", "lnR"), ("top", "lnT"), ("bottom", "lnB")):
                            c = color(pr.find("a:" + key, NS))
                            if c:
                                css.append("border-%s:1px solid %s" % (side, c))
                        if color(pr):
                            css.append("background:" + color(pr))
                    cells.append({"css": ";".join(css), "paras": parse_paras(tc.find("a:txBody", NS), rels, recolor)})
                e["rows"].append({"h": int(tr.get("h")) / EMU, "cells": cells})
            e["h"] = sum(r["h"] for r in e["rows"])
        els.append(e)
    return bg, els


def assign_roles(els):
    """標出頁首小標（eyebrow）、標題（title）與頁尾來源列（source）；兩種來源簡報的形狀命名不同，所以看結構判斷。"""
    title = next((e for e in els if e["ph"] == "title" and "paras" in e), None)
    for e in els:
        if "paras" not in e or e.get("fill"):
            continue
        if e is title:
            e["role"] = "title"
            runs = [r for p in e["paras"] for r in p["runs"]]
            for r in runs:  # 標題裡個別字被縮小的（例如手動改過的「AI」）拉回同一字級
                r["z"] = max(x["z"] for x in runs)
        elif title is not None and e["y"] + e["h"] <= title["y"] + 2:
            e["role"] = "eyebrow"
        elif e["y"] > 560 and re.match(r"(資料|案例)?來源[：:]|資料分級[：:]|依 OWASP|Skill 範例為|Skill 設計為", plain(e)):
            e["role"] = "source"
    return title


# ───────────────────────── patches.py 用的工具 ─────────────────────────

def inside(inner, outer, tol=3):
    cx, cy = inner["x"] + inner["w"] / 2, inner["y"] + inner["h"] / 2
    return (outer["x"] - tol <= cx <= outer["x"] + outer["w"] + tol
            and outer["y"] - tol <= cy <= outer["y"] + outer["h"] + tol
            and inner["h"] <= outer["h"] + tol)


def find(els, pat):
    """內文符合正規表示式的文字元素"""
    out = [e for e in els if "paras" in e and re.search(pat, plain(e))]
    assert out, "找不到文字：" + pat
    return out


def box_of(els, e):
    """包住 e 的最小色塊（卡片）"""
    cands = [c for c in els if c is not e and "paras" not in c and c["tag"] == "sp"
             and (c.get("fill") or c.get("stroke")) and inside(e, c) and c["w"] * c["h"] > e["w"] * e["h"]]
    return min(cands, key=lambda c: c["w"] * c["h"]) if cands else None


def members(els, box):
    return [o for o in els if o is not box and inside(o, box) and o["w"] * o["h"] < box["w"] * box["h"]]


def remove(els, *targets):
    for t in targets:
        els[:] = [e for e in els if e is not t]


def rm_card(els, pat):
    """刪掉含指定文字的整張卡片（色塊與裡面的東西）"""
    for e in find(els, pat):
        box = box_of(els, e) or e
        top, bottom = box["y"], box["y"] + box["h"]
        remove(els, e, box, *members(els, box))
        # 刪掉的是整列時，下方內容往上遞補
        rest = [o for o in els if not o["role"]]
        below = [o for o in rest if o["y"] >= bottom - 1]
        if below and not any(o["y"] < bottom - 1 and o["y"] + o["h"] > top + 1 for o in rest):
            dy = min(o["y"] for o in below) - top
            for o in below:
                o["y"] -= dy


def _span(para, start, end):
    """把段落的 run 在 [start, end) 邊界切開，回傳涵蓋範圍的 run 索引"""
    runs, pos, i, cover = para["runs"], 0, 0, []
    while i < len(runs):
        r = runs[i]
        a, b = pos, pos + len(r["text"])
        for at in (start, end):
            if a < at < b:
                left, right = dict(r), dict(r)
                left["text"], right["text"] = r["text"][:at - a], r["text"][at - a:]
                runs[i:i + 1] = [left, right]
                r, b = left, at
        if a >= start and b <= end and b > a:
            cover.append(i)
        pos, i = b, i + 1
    return cover


def mark(e, phrase, col=RUST):
    """重點標註：把 phrase 改成指定顏色的粗體"""
    hit = False
    for para in e["paras"]:
        at = ptext(para).find(phrase)
        if at >= 0:
            for i in _span(para, at, at + len(phrase)):
                para["runs"][i]["color"], para["runs"][i]["b"] = col, True
            hit = True
    assert hit, "找不到要標註的文字：" + phrase


def sub(e, old, new):
    """取代文字（可跨 run）"""
    hit = False
    for para in e["paras"]:
        at = ptext(para).find(old)
        if at >= 0:
            idx = _span(para, at, at + len(old))
            para["runs"][idx[0]]["text"] = new
            for i in idx[1:]:
                para["runs"][i]["text"] = ""
            hit = True
    assert hit, "找不到要取代的文字：" + old


def cut(e, pat):
    """刪掉符合正規表示式的文字；刪完整段都空了就拿掉整段"""
    hit = False
    for para in list(e["paras"]):
        m = re.search(pat, ptext(para))
        if not m:
            continue
        hit = True
        for i in _span(para, m.start(), m.end()):
            para["runs"][i]["text"] = ""
        if not ptext(para).strip() and len(e["paras"]) > 1:
            e["paras"].remove(para)
    assert hit, "找不到要刪的文字：" + pat


def set_text(para, text):
    """整段換字，沿用最後一個 run 的樣式"""
    last = dict(para["runs"][-1])
    last["text"] = text
    para["runs"] = [last]


def add_para(e, text, like=-1):
    para = copy.deepcopy(e["paras"][like])
    set_text(para, text)
    e["paras"].append(para)
    e["free"] = True  # 內容變長了，交給排版引擎重新撐開
    return para


def set_list(e, items, gap=8):
    """把文字框改成編號條列（1. 2. 3.），沿用原本的字級與顏色"""
    base = e["paras"][0]
    e["paras"], e["free"] = [], True
    for k, item in enumerate(items, 1):
        para = copy.deepcopy(base)
        set_text(para, "%d. %s" % (k, item))
        para["sb"], para["hang"], para["bu"] = (gap if k > 1 else 0), 1.3, None
        e["paras"].append(para)


def merge_heads(els, sep=". ", keep_num=True):
    """卡片標頭「1／角色背景」兩行併成一行「1. 角色背景」，卡片內其餘內容與下方各列跟著上移，讓出空間。
    keep_num=False 時直接拿掉那一行編號。"""
    heads = [e for e in els if not e["role"] and "paras" in e and len(e["paras"]) >= 2
             and re.fullmatch(r"\d{1,2}", ptext(e["paras"][0]).strip())]
    assert heads, "找不到數字標頭"
    rows = {}
    for e in heads:
        rows.setdefault(round(e["y"]), []).append(e)
    for key in sorted(rows):
        touched, bottom, delta = [], 0, 0
        for e in rows[key]:
            p0, p1 = e["paras"][0], e["paras"][1]
            delta = p0["lh"] + p1["sb"]
            if keep_num:
                num = dict(p1["runs"][0])
                num["text"], num["color"] = ptext(p0).strip() + sep, p0["runs"][0]["color"]
                p1["runs"].insert(0, num)
            p1["sb"] = 0
            del e["paras"][0]
            e["h"] -= delta
            e["free"] = True
            box = box_of(els, e)
            inner = members(els, box)
            touched += [e, box] + inner
            bottom = max(bottom, box["y"] + box["h"])
            for o in inner:
                if o is not e and o["y"] > e["y"] + 5:
                    o["y"] -= delta
            box["h"] -= delta
        for o in els:
            if not o["role"] and not any(o is t for t in touched) and o["y"] >= bottom - 1:
                o["y"] -= delta


def stack(els, boxes):
    """把同一張卡片裡上下相疊的幾個單行文字框併成一個（某一行折行時，其他行才不會被拉開）"""
    boxes = sorted(boxes, key=lambda e: e["y"])
    first, prev = boxes[0], boxes[0]
    for nxt in boxes[1:]:
        nxt["paras"][0]["sb"] = max(6, nxt["y"] - (prev["y"] + prev["h"]))
        first["paras"] += nxt["paras"]
        first["w"] = max(first["w"], nxt["w"])
        prev = nxt
        remove(els, nxt)
    first["h"], first["free"] = prev["y"] + prev["h"] - first["y"], True


def shrink_row(els, texts, new_h):
    """文字被濃縮後，把這一列的文字框（與所在卡片）改矮，下方內容往上遞補，空間交還給排版引擎"""
    delta = texts[0]["h"] - new_h
    bottom = max((box_of(els, e) or e)["y"] + (box_of(els, e) or e)["h"] for e in texts)
    for e in texts:
        box = box_of(els, e)
        e["h"] = new_h
        if box:
            box["h"] -= delta
    for o in els:
        if not o["role"] and o["y"] >= bottom - 1:
            o["y"] -= delta


def recolor_pills(els, fill, ink):
    """膠囊標籤換色（底色與字色）"""
    pills = [e for e in els if e["tag"] == "sp" and e.get("prst") == "roundRect" and e.get("adj", 0) >= 40000
             and e.get("fill") and e["h"] > 8]
    assert pills, "找不到膠囊標籤"
    for p in pills:
        p["fill"] = fill
        for t in els:
            if "paras" in t and (t is p or inside(t, p)):
                for r in (r for para in t["paras"] for r in para["runs"]):
                    r["color"] = ink


def respread(els, boxes, x0=85.3, x1=1194.4, gap=14.6):
    """把一列卡片重新等分排滿寬度（刪掉其中一張之後用）"""
    boxes = sorted(boxes, key=lambda b: b["x"])
    w = (x1 - x0 - gap * (len(boxes) - 1)) / len(boxes)
    for i, box in enumerate(boxes):
        nx, ratio = x0 + i * (w + gap), w / box["w"]
        for o in members(els, box):
            o["x"] = nx + (o["x"] - box["x"]) * ratio
            if "paras" in o:
                o["w"] *= ratio
        box["x"], box["w"] = nx, w


# ───────────────────────── 版面分析 ─────────────────────────

def analyze(els):
    """內容頁：切帶、算每個文字框可以長到哪裡。回傳給 <section> 的屬性字串。"""
    for e in els:
        e["y"] += 8 if e["role"] == "source" else -SHIFT_UP
        # 版心加寬：所有元素的水平位置與寬度等比放大
        fixed = e["tag"] == "pic" or (e.get("prst") == "roundRect" and e.get("adj", 0) >= 40000 and abs(e["w"] - e["h"]) <= 2)
        mid = LEFT + (e["x"] + e["w"] / 2 - 85.3) * WIDEN
        if e["tag"] == "graphicFrame":
            e["cols"] = [c * WIDEN for c in e["cols"]]
        if not fixed:
            e["w"] *= WIDEN
        e["x"] = mid - e["w"] / 2
        if "paras" in e and not e.get("fill") and e["x"] + e["w"] > RIGHT:
            e["w"] = RIGHT - e["x"]
    body = [e for e in els if not e["role"]]
    if not body:
        return ""
    # 標題原本折成兩行、加寬後排得進一行：下方內容往上遞補
    title = next(e for e in els if e["role"] == "title")
    est = sum(1 if ord(c) > 0x2E7F else 0.56 for c in plain(title)) * title["paras"][0]["runs"][0]["z"]
    if title["h"] > 70 and est < RIGHT - LEFT - 30:
        for e in body:
            e["y"] -= title["h"] - 51
    src = [e for e in els if e["role"] == "source"]
    limit = min(e["y"] for e in src) - 10 if src else BOTTOM_FREE

    def is_in(e, group):
        return any(e is o for o in group)

    # 膠囊標籤、圓形編號與疊在上面的字：不跟著拉高，只依字級等比放大
    pills = [e for e in body if e["tag"] == "sp" and e.get("prst") == "roundRect" and e["adj"] >= 40000
             and e.get("fill") and e["h"] > 8]
    for p in pills:
        p["kind"] = "r" if abs(p["w"] - p["h"]) <= 2 else "s"
        if "paras" in p:
            z = p["paras"][0]["runs"][0]["z"]
            p["key"], p["m"] = "s%s" % n(z), gain(z)
    for e in body:
        if "paras" in e and not e.get("fill") and "kind" not in e:
            p = next((p for p in pills if inside(e, p)), None)
            if p is not None:
                z = e["paras"][0]["runs"][0]["z"]
                e["kind"], e["key"], e["m"] = p["kind"], "s%s" % n(z), gain(z)
                p["key"], p["m"] = e["key"], e["m"]

    def holds(c, o):
        return o is not c and inside(o, c) and o["w"] * o["h"] < c["w"] * c["h"]

    conts = [c for c in body if c["tag"] == "sp" and c.get("prst") in ("rect", "roundRect") and "paras" not in c
             and "kind" not in c and (c.get("fill") or c.get("stroke")) and any(holds(c, o) for o in body)]
    for c in conts:
        held = [o for o in body if holds(c, o)]
        if (len(held) == 1 and "paras" in held[0] and not held[0].get("fill") and not held[0].get("free")
                and abs(held[0]["y"] + held[0]["h"] / 2 - c["y"] - c["h"] / 2) <= 6):
            held[0]["anchor"] = "ctr"  # 色塊裡只有一個原本就置中的文字框：拉高後維持垂直置中
    for e in body:
        if "pad" in e and e.get("fill") and "kind" not in e and not any(holds(e, o) for o in body):
            e["anchor"] = "ctr"
            e["pad"][0] = e["pad"][2] = min(e["pad"][0], e["pad"][2])

    leaves = [e for e in body if not is_in(e, conts)]
    bands = []
    for a, b in sorted((e["y"], e["y"] + e["h"]) for e in leaves):
        if bands and (a < bands[-1][1] - 3 or (b - a <= 2 and a <= bands[-1][1] + 1)):  # 貼著上一帶下緣的細線併進去
            bands[-1][1] = max(bands[-1][1], b)
        else:
            bands.append([a, b])

    def band_of(e):
        mid = e["y"] + e["h"] / 2
        return next((i for i, (a, b) in enumerate(bands) if a - 0.5 <= mid <= b + 0.5), -1)

    tail = max(0, max(e["y"] + e["h"] for e in body) - bands[-1][1])
    for k, e in enumerate(body):
        e["band"] = -1 if is_in(e, conts) else band_of(e)
        if "kind" not in e:
            flat = e["tag"] in ("pic", "cxnSp") or (e.get("prst") == "line" and e["w"] >= e["h"])
            e["kind"] = "p" if flat else "a"
        if not (e["tag"] == "graphicFrame" or ("paras" in e and e["kind"] == "a")):
            continue
        i = e["band"]
        b = bands[i][1]
        nxt = bands[i + 1][0] if i + 1 < len(bands) else b + tail  # 最後一帶之後只剩卡片下緣的留白
        mine = []
        if "pad" in e and e.get("fill"):
            e["room"] = (e["h"], -(e["pad"][0] + e["pad"][2]))
        else:
            lims = [limit]
            mine = [c for c in conts if holds(c, e)]
            if mine:
                box = min(mine, key=lambda c: c["w"] * c["h"])
                lims.append(box["y"] + box["h"] - max(4, min(e["y"] - box["y"], e["x"] - box["x"], 18)))
            for o in body:
                if o is e or is_in(o, mine):
                    continue
                overlap = min(o["x"] + o["w"], e["x"] + e["w"]) - max(o["x"], e["x"])
                if o["y"] >= e["y"] + 6 and overlap > 4:
                    lims.append(o["y"] - 3)
            lim = min(lims)
            e["room"] = (max(8, min(lim, b) - e["y"]), max(0, min(lim, nxt) - b))
        if e["tag"] == "graphicFrame":
            e["key"] = "T%d" % k
        else:
            e["key"] = "%s|%d" % (n(e["paras"][0]["runs"][0]["z"]), round(e["w"]))
            if (i == 0 and not mine and not e.get("fill") and not e.get("boost") and not e.get("free")
                    and len([o for o in leaves if band_of(o) == 0]) == 1):
                e["key"] = "L:" + e["key"]  # 標題下方的引言：內文都排好之後，有剩餘空間才放大
    return ' data-bands="%s" data-lim="%s" data-tail="%s"' % (
        ";".join("%s,%s" % (n(a), n(b)) for a, b in bands), n(limit), n(tail))


# ───────────────────────── 輸出 HTML ─────────────────────────

def run_cls(r):
    css = ["--z:%s;--m:%s" % (n(r["z"]), gain(r["z"]))]
    if r["b"]:
        css.append("font-weight:700")
    if r["i"]:
        css.append("font-style:italic")
    if r["u"]:
        css.append("text-decoration:underline")
    if r["spc"]:
        css.append("letter-spacing:%spx" % n(r["spc"]))
    if r["color"]:
        css.append("color:" + r["color"])
    if r["serif"]:
        css.append('font-family:"Noto Serif TC","PMingLiU",serif')
    return RUNS.cls(";".join(css))


def render_paras(paras):
    out = []
    for para in paras:
        z0 = para["runs"][0]["z"]
        css = ["--lh:%s;--z:%s;--m:%s" % (n(para["lh"]), n(z0), gain(z0))]
        if para["sb"]:
            css.append("--sb:%s" % n(para["sb"]))
        if para["algn"]:
            css.append("text-align:" + para["algn"])
        bullet = ""
        if para["bu"]:
            char, ind = para["bu"]
            css.append("padding-left:%spx;text-indent:-%spx" % (n(ind), n(ind)))
            bullet = '<span class="bu %s" style="width:%spx">%s</span>' % (
                run_cls(para["runs"][0]), n(ind), html.escape(char))
        if para["hang"]:
            css.append("padding-left:%sem;text-indent:-%sem" % (para["hang"], para["hang"]))
        body = []
        for r in para["runs"]:
            if not r["text"]:
                continue
            span = "<br>" if r["text"] == "\n" else '<span class="%s">%s</span>' % (run_cls(r), html.escape(r["text"]))
            if r["href"]:
                span = '<a href="%s" target="_blank" rel="noopener">%s</a>' % (html.escape(r["href"]), span)
            body.append(span)
        style = ' style="--x:%s"' % para["x"] if para["x"] else ""
        out.append('<p class="%s"%s>%s%s</p>' % (PARAS.cls(";".join(css)), style, bullet, "".join(body)))
    return "".join(out)


def data_attrs(e, top, height):
    """給瀏覽器端排版用的資料：類型、原始位置、所在的帶、字級群組、可用高度"""
    if "band" not in e:
        return ""
    vals = [e["kind"], n(top), n(height), str(e["band"]), str(e.get("m", 0))]
    if e["kind"] == "r":
        vals += [n(e["x"]), n(e["w"])]
    out = ' data-g="%s"' % ",".join(vals)
    if "key" in e:
        out += ' data-q="%s"' % e["key"]
    if "room" in e:
        out += ' data-c="%s,%s%s"' % (n(e["room"][0]), n(e["room"][1]), ",1" if e.get("free") else "")
    return out


def render(e, images, static):
    geo = "left:%spx;top:%spx;width:%spx;height:%spx" % (n(e["x"]), n(e["y"]), n(e["w"]), n(e["h"]))
    if e["tag"] == "pic":
        return '<i class="e %s" style="%s"%s></i>' % (images[e["img"]], geo, data_attrs(e, e["y"], e["h"]))
    if e["tag"] == "graphicFrame":
        cols = "".join('<col style="width:%spx">' % n(w) for w in e["cols"])
        rows = "".join('<tr data-h="%s" style="height:%spx">%s</tr>' % (n(r["h"]), n(r["h"]), "".join(
            '<td style="%s">%s</td>' % (c["css"], render_paras(c["paras"])) for c in r["cells"])) for r in e["rows"])
        style = "left:%spx;top:%spx;width:%spx" % (n(e["x"]), n(e["y"]), n(sum(e["cols"])))
        if e.get("boost"):
            style += ";--x:%s" % e["boost"]
        return '<table class="e tbl" style="%s"%s><colgroup>%s</colgroup>%s</table>' % (
            style, data_attrs(e, e["y"], e["h"]), cols, rows)
    stroke = e.get("stroke")
    if e["tag"] == "cxnSp" or e["prst"] == "line":
        sw, c = e["sw"] or 1, stroke or "#000"
        if not e.get("arrow"):
            if e["w"] < 0.5:
                return '<i class="e" style="left:%spx;top:%spx;width:%spx;height:%spx;background:%s"%s></i>' % (
                    n(e["x"] - sw / 2), n(e["y"]), n(sw), n(e["h"]), c, data_attrs(e, e["y"], e["h"]))
            return '<i class="e" style="left:%spx;top:%spx;width:%spx;height:%spx;background:%s"%s></i>' % (
                n(e["x"]), n(e["y"] - sw / 2), n(e["w"]), n(sw), c, data_attrs(e, e["y"] - sw / 2, sw))
        a, w = sw * 3.2, e["w"]  # a：箭頭大小
        return ('<svg class="e" style="left:%spx;top:%spx;width:%spx;height:%spx" viewBox="0 0 %s %s"%s>'
                '<path d="M0 %sH%s" stroke="%s" stroke-width="%s"/><path d="M%s %sl-%s-%sv%sz" fill="%s"/></svg>'
                % (n(e["x"]), n(e["y"] - a), n(w), n(a * 2), n(w), n(a * 2), data_attrs(e, e["y"] - a, a * 2),
                   n(a), n(w - a), c, n(sw), n(w), n(a), n(a * 1.3), n(a * .75), n(a * 1.5), c))
    css = [geo]
    if e.get("fill"):
        css.append("background:" + e["fill"])
    if stroke:
        css.append("border:%spx solid %s" % (n(1 if e["sw"] < 1.5 else e["sw"]), stroke))
    if e["prst"] == "roundRect":
        css.append("border-radius:%spx" % ("999" if e["adj"] >= 40000 else n(e["adj"] / 100000 * min(e["w"], e["h"]))))
    if "paras" not in e:
        return '<i class="e" style="%s"%s></i>' % (";".join(css), data_attrs(e, e["y"], e["h"]))
    if any(v > 0.5 for v in e["pad"]):
        css.append("padding:%spx %spx %spx %spx" % tuple(n(v) for v in e["pad"]))
    fixed = {"eyebrow": EYEBROW_K, "source": SOURCE_K, "title": 1}.get(e["role"])
    if fixed is not None:
        # 頁首、頁尾用固定倍率，不參與自動放大
        m = gain(e["paras"][0]["runs"][0]["z"]) or 1
        css.append("--t:1;--x:%.3f" % ((fixed - 1) / m))
    elif e.get("boost"):
        css.append(("--t:1;" if static else "") + "--x:%s" % e["boost"])
    cls = "e tb" + (" ac" if e["anchor"] == "ctr" else "") + (" nw" if e.get("kind") in ("s", "r") else "")
    return '<div class="%s" style="%s"%s><div class="ti">%s</div></div>' % (
        cls, ";".join(css), data_attrs(e, e["y"], e["h"]), render_paras(e["paras"]))


def build(src, out):
    import patches

    z = zipfile.ZipFile(src)
    order = etree.fromstring(z.read("ppt/presentation.xml")).findall("p:sldIdLst/p:sldId", NS)
    prels = {r.get("Id"): r.get("Target") for r in etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))}
    files = ["ppt/" + prels[s.get(RID)] for s in order]

    # 互動演練頁與手工重新設計的頁：都是以 <!--@原頁碼--> 標示、整頁取代 PPT 的同一頁
    frag = "".join(re.sub(r"^.*?(?=<!--@\d+-->)", "", (HERE / f).read_text(encoding="utf-8"), flags=re.S)
                   for f in ("interactive.html", "custom.html"))  # 各檔第一個標記之前是說明文字，丟掉
    frag = re.sub(r"@@IMG:(\w+)@@", lambda m: "data:image/jpeg;base64," + base64.b64encode(
        (HERE / "images" / (m.group(1) + ".jpg")).read_bytes()).decode(), frag)
    inter = {int(m.group(1)): m.group(2).strip()
             for m in re.finditer(r"<!--@(\d+)-->(.*?)(?=<!--@\d+-->|\Z)", frag, re.S)}

    images, img_css, sections, pages = {}, [], [], {}
    for i, f in enumerate(files, 1):
        if i in patches.DROP:
            continue
        pages[i] = len(sections) + 1
        rel_path = f.replace("slides/", "slides/_rels/") + ".rels"
        rels = {}
        if rel_path in z.namelist():
            for r in etree.fromstring(z.read(rel_path)):
                rels[r.get("Id")] = r.get("Target")
        recolor = {k.upper(): v for k, v in patches.RECOLOR.get(i, {}).items()}
        bg, els = parse_slide(z.read(f), rels, recolor)
        t = assign_roles(els)
        if i not in inter:
            try:
                patches.apply(i, els, sys.modules[__name__])
            except AssertionError as err:
                raise SystemExit("第 %d 頁的調整失敗：%s" % (i, err))
        title = plain(t).replace("\n", " ") if t else ""
        if not title:
            texts = sorted((e for e in els if "paras" in e), key=lambda e: -e["h"])
            title = plain(texts[0]).replace("\n", " ") if texts else "第 %d 頁" % i
        if i in inter:
            sections.append(inter[i].replace("<section ", '<section data-src="%d" data-t="%s" ' % (i, html.escape(title)), 1))
            continue
        for e in els:
            if e["tag"] == "pic" and e["img"] not in images:
                data = z.read("ppt/" + e["img"][3:])
                images[e["img"]] = "im%d" % len(images)
                img_css.append(".%s{background:url(data:image/png;base64,%s) center/100%% 100%% no-repeat}"
                               % (images[e["img"]], base64.b64encode(data).decode()))
        static = bg.upper() != "#F5F2EC" or t is None
        attrs = "" if static else analyze(els)
        if attrs and i in patches.TMAX:
            attrs += ' data-tmax="%s"' % patches.TMAX[i]
        cls = {"#15233A": " bn", "#A3500C": " br"}.get(bg.upper(), "")
        sections.append('<section class="s%s" data-src="%d" data-t="%s"%s>%s</section>' % (
            cls, i, html.escape(title), attrs, "".join(render(e, images, static) for e in els)))

    tpl = (HERE / "template.html").read_text(encoding="utf-8")
    page = (tpl.replace("/*GEN_CSS*/", RUNS.css() + PARAS.css() + "".join(img_css))
            .replace("<!--SLIDES-->", "\n".join(sections))
            .replace("/*QR_LIB*/", (HERE / "qrcode.min.js").read_text(encoding="utf-8")))
    Path(out).write_text(page, encoding="utf-8")
    print("頁數: %d（原 %d 頁，刪除原第 %s 頁），檔案: %.0f KB" % (
        len(sections), len(files), "、".join(map(str, sorted(patches.DROP))), len(page.encode("utf-8")) / 1024))
    print("互動／重新設計頁: " + ", ".join("原 p.%d → 新 p.%d" % (k, pages[k]) for k in sorted(inter)))
    return pages


if __name__ == "__main__":
    sys.path.insert(0, str(HERE))
    build(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else str(HERE.parent / "index.html"))
