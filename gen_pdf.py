#!/usr/bin/env python3
"""Generate 渋一プロジェクト概要.pdf - A4 with HeiseiKakuGo-W5 CID font"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
import os

# Register CID font
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))
FONT = 'HeiseiKakuGo-W5'

# Colors
RED = HexColor('#CC0000')
GRAY = HexColor('#888888')
GOLD = HexColor('#D4A845')
LIGHT_GRAY = HexColor('#AAAAAA')
CARD_BG = HexColor('#F5F5F5')
HEADER_BG = HexColor('#1A1A1A')

# Page setup
PAGE_W, PAGE_H = A4
MARGIN = 18 * mm
CONTENT_W = PAGE_W - 2 * MARGIN

# Font sizes
TITLE_SIZE = 28
HEADER_SIZE = 11
BODY_SIZE = 8.5
SMALL_SIZE = 7.5
LABEL_SIZE = 7

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "\u6e0b\u4e00\u30d7\u30ed\u30b8\u30a7\u30af\u30c8\u6982\u8981.pdf")
c = canvas.Canvas(OUT, pagesize=A4)
c.setTitle("\u6e0b\u4e00\u30d7\u30ed\u30b8\u30a7\u30af\u30c8\u6982\u8981")


def draw_header_bar(y, text, color=HEADER_BG, text_color=white):
    bar_h = 7 * mm
    c.setFillColor(color)
    c.rect(MARGIN, y - bar_h, CONTENT_W, bar_h, fill=1, stroke=0)
    c.setFillColor(text_color)
    c.setFont(FONT, HEADER_SIZE)
    c.drawString(MARGIN + 3 * mm, y - bar_h + 2 * mm, text)
    return y - bar_h - 2 * mm


def draw_wrapped_text(x, y, text, size=BODY_SIZE, color=black, max_width=None, line_height=None):
    if max_width is None:
        max_width = CONTENT_W - (x - MARGIN)
    if line_height is None:
        line_height = size * 0.5 * mm + 1 * mm
    c.setFillColor(color)
    c.setFont(FONT, size)
    char_width = size * 0.35 * mm
    chars_per_line = int(max_width / char_width)
    lines = []
    for paragraph in text.split('\n'):
        while len(paragraph) > chars_per_line:
            lines.append(paragraph[:chars_per_line])
            paragraph = paragraph[chars_per_line:]
        lines.append(paragraph)
    for line in lines:
        c.drawString(x, y, line)
        y -= line_height
    return y


def new_page():
    c.showPage()


# ═══════════════════════════════════════════════════════
# PAGE 1: TITLE
# ═══════════════════════════════════════════════════════
c.setFillColor(RED)
c.rect(0, PAGE_H - 3 * mm, PAGE_W, 3 * mm, fill=1, stroke=0)

y = PAGE_H - 45 * mm
c.setFillColor(black)
c.setFont(FONT, TITLE_SIZE)
title = "\u6e0b\u4e00\uff08\u30b7\u30d6\u30ef\u30f3\uff09"
tw = c.stringWidth(title, FONT, TITLE_SIZE)
c.drawString((PAGE_W - tw) / 2, y, title)

y -= 15 * mm
c.setFont(FONT, 14)
c.setFillColor(GRAY)
sub1 = "\u6e0b\u8c37\u767a\u30e1\u30f3\u30baIP\u5275\u51fa\u30d7\u30ed\u30b8\u30a7\u30af\u30c8"
tw = c.stringWidth(sub1, FONT, 14)
c.drawString((PAGE_W - tw) / 2, y, sub1)

y -= 10 * mm
c.setFont(FONT, 10)
c.setFillColor(LIGHT_GRAY)
sub2 = "Shibuya Ikemen No.1 Project  \u00d7  \u30b9\u30c8\u30fc\u30ea\u30fc\u30b3\u30f3\u30c6\u30f3\u30c4"
tw = c.stringWidth(sub2, FONT, 10)
c.drawString((PAGE_W - tw) / 2, y, sub2)

y -= 15 * mm
c.setFillColor(GOLD)
c.setFont(FONT, 12)
sub3 = "\u4f01\u753b\u66f8  \u30d7\u30ed\u30b8\u30a7\u30af\u30c8\u6982\u8981"
tw = c.stringWidth(sub3, FONT, 12)
c.drawString((PAGE_W - tw) / 2, y, sub3)

y -= 15 * mm
c.setStrokeColor(RED)
c.setLineWidth(1)
c.line(PAGE_W / 2 - 40 * mm, y, PAGE_W / 2 + 40 * mm, y)

y -= 15 * mm
c.setFillColor(GRAY)
c.setFont(FONT, 8)
conf = "CONFIDENTIAL  \u00a9  HJ Inc. All Rights Reserved."
tw = c.stringWidth(conf, FONT, 8)
c.drawString((PAGE_W - tw) / 2, y, conf)


# ═══════════════════════════════════════════════════════
# PAGE 2: MISSION + CONCEPT
# ═══════════════════════════════════════════════════════
new_page()
y = PAGE_H - MARGIN
c.setFillColor(RED)
c.rect(MARGIN - 3 * mm, MARGIN, 2 * mm, PAGE_H - 2 * MARGIN, fill=1, stroke=0)

y = draw_header_bar(y, "MISSION", color=RED)
y -= 5 * mm
y = draw_wrapped_text(MARGIN, y,
    "Z\u30fb\u03b1\u4e16\u4ee3\u5973\u6027\u304c\u300c\u6700\u521d\u306b\u63a8\u3057\u305f\u7537\u6027\u30b9\u30bf\u30fc\u300d\u3092\u6bce\u5e74\u751f\u307f\u51fa\u3057\u7d9a\u3051\u308b\u3001\n\u65e5\u672c\u6700\u5927\u306e\u30e1\u30f3\u30baIP\u5275\u51fa\u88c5\u7f6e\u3002",
    size=11, color=black)

y -= 5 * mm
pillars = [
    ("No.1 \u30dd\u30b8\u30b7\u30e7\u30f3\u78ba\u7acb", "\u30e1\u30f3\u30ba\u30a2\u30a4\u30c9\u30eb\u5e02\u5834\u306b\u304a\u3051\u308b\u5727\u5012\u7684\u30b7\u30a7\u30a2\u3092\u7372\u5f97", RED),
    ("\u5727\u5012\u7684\u306a\u8a71\u984c\u6027", "SNS\u30d0\u30ba\u3001\u30e1\u30c7\u30a3\u30a2\u9732\u51fa\u3001\u30c8\u30ec\u30f3\u30c9\u5165\u308a\u3092\u5e38\u614b\u5316", GOLD),
    ("\u6700\u5f37\u306e\u30d5\u30a1\u30f3\u30c0\u30e0", "\u71b1\u91cf\u306e\u9ad8\u3044\u30d5\u30a1\u30f3\u30b3\u30df\u30e5\u30cb\u30c6\u30a3\u3068\u7d99\u7d9a\u7684\u306a\u7d4c\u6e08\u570f\u3092\u5275\u51fa", HexColor('#1E6FBF')),
]
for title, desc, accent in pillars:
    c.setFillColor(CARD_BG)
    c.rect(MARGIN, y - 14 * mm, CONTENT_W, 14 * mm, fill=1, stroke=0)
    c.setFillColor(accent)
    c.rect(MARGIN, y, CONTENT_W, 1 * mm, fill=1, stroke=0)
    c.setFont(FONT, 9)
    c.drawString(MARGIN + 3 * mm, y - 5 * mm, title)
    c.setFillColor(GRAY)
    c.setFont(FONT, BODY_SIZE)
    c.drawString(MARGIN + 3 * mm, y - 11 * mm, desc)
    y -= 17 * mm

y -= 3 * mm
c.setFillColor(HexColor('#FFF0F0'))
c.rect(MARGIN, y - 8 * mm, CONTENT_W, 8 * mm, fill=1, stroke=0)
c.setFillColor(RED)
c.setFont(FONT, 9)
target = "\u76ee\u6a19:  2028\u5e74  \u30e1\u30f3\u30ba\u82b8\u80fd\u90e8\u9580\uff11\u4f4d  \uff1d  100\u5104\u5186IP\u306e\u5b9f\u73fe"
tw = c.stringWidth(target, FONT, 9)
c.drawString((PAGE_W - tw) / 2, y - 5.5 * mm, target)

y -= 15 * mm
y = draw_header_bar(y, "CONCEPT", color=RED)
y -= 5 * mm
c.setFillColor(black)
c.setFont(FONT, 10)
c.drawString(MARGIN, y, "\u300c\u653e\u8ab2\u5f8c\u6f14\u6280\u6d3e\u30af\u30e9\u30d6\u300d\u2500\u2500 \u30b9\u30c8\u30fc\u30ea\u30fc \u00d7 \u30a2\u30a4\u30c9\u30eb")
y -= 6 * mm
y = draw_wrapped_text(MARGIN, y,
    "\u6e0b\u8c37\u306e10\u306e\u67b6\u7a7a\u9ad8\u6821\u3092\u821e\u53f0\u306b\u3001\u5404\u6821\u304c\u30e1\u30f3\u30ba\u30dc\u30fc\u30ab\u30eb\u30e6\u30cb\u30c3\u30c8\u3092\u7d50\u6210\u3002\n\u30b7\u30e7\u30fc\u30c8\u30c9\u30e9\u30de\u3067\u63cf\u304f\u9752\u6625\u30b9\u30c8\u30fc\u30ea\u30fc\u3068\u3001\u30ea\u30a2\u30eb\u306e\u30d1\u30d5\u30a9\u30fc\u30de\u30f3\u30b9\u304c\u4ea4\u5dee\u3059\u308b\u65b0\u4e16\u4ee3\u30a8\u30f3\u30bf\u30e1\u3002",
    size=BODY_SIZE, color=GRAY)

y -= 3 * mm
col_w = (CONTENT_W - 4 * mm) / 2
axes = [
    ("\u89aa\u8fd1\u611f\u8ef8\uff08\u30b3\u30e0\u30c9\u30c3\u30c8\u7684\uff09", [
        "\u65e5\u5e38\u30fb\u30d0\u30e9\u30a8\u30c6\u30a3\u4f01\u753b",
        "\u30aa\u30d5\u30b7\u30e7\u30c3\u30c8\u30fb\u7d20\u306e\u8868\u60c5",
        "\u30e1\u30f3\u30d0\u30fc\u9593\u306e\u95a2\u4fc2\u6027",
        "\u5207\u308a\u629c\u304d\u62e1\u6563\u30fb\u30b7\u30e7\u30fc\u30c8\u52d5\u753b",
    ], GOLD),
    ("\u30a2\u30a4\u30c9\u30eb\u8ef8\uff08\u30b9\u30bf\u30fc\u6027\u78ba\u7acb\uff09", [
        "\u697d\u66f2\u30fb\u30c0\u30f3\u30b9\u30d1\u30d5\u30a9\u30fc\u30de\u30f3\u30b9",
        "\u30c9\u30e9\u30de\u30fb\u30b9\u30c8\u30fc\u30ea\u30fc\u3078\u306e\u611f\u60c5\u79fb\u5165",
        "\u30d3\u30b8\u30e5\u30a2\u30eb\u91cd\u8996\u306e\u9ad8\u54c1\u8cea\u30b3\u30f3\u30c6\u30f3\u30c4",
        "\u30d5\u30a1\u30f3\u4ea4\u6d41\u30a4\u30d9\u30f3\u30c8\u30fb\u30e9\u30a4\u30d6",
    ], RED),
]
for i, (title, items, accent) in enumerate(axes):
    bx = MARGIN + i * (col_w + 4 * mm)
    card_h = 28 * mm
    c.setFillColor(CARD_BG)
    c.rect(bx, y - card_h, col_w, card_h, fill=1, stroke=0)
    c.setFillColor(accent)
    c.rect(bx, y, col_w, 0.8 * mm, fill=1, stroke=0)
    c.setFont(FONT, 8)
    c.drawString(bx + 2 * mm, y - 5 * mm, title)
    iy = y - 10 * mm
    c.setFillColor(GRAY)
    c.setFont(FONT, SMALL_SIZE)
    for item in items:
        c.drawString(bx + 4 * mm, iy, "\u25cf  " + item)
        iy -= 4.5 * mm


# ═══════════════════════════════════════════════════════
# PAGE 3: WORLD BUILDING
# ═══════════════════════════════════════════════════════
new_page()
y = PAGE_H - MARGIN
c.setFillColor(RED)
c.rect(MARGIN - 3 * mm, MARGIN, 2 * mm, PAGE_H - 2 * MARGIN, fill=1, stroke=0)

y = draw_header_bar(y, "WORLD BUILDING \u2500 \u6e0b\u8c3710\u6821", color=RED)

schools = [
    ("\u6e0b\u4e00", "\u5b87\u7530\u5ddd\u753a", "\u30b0\u30e9\u30d5\u30a3\u30c6\u30a3", "#CC0000", "\u308b\u3044\uff08\u4e3b\u4eba\u516c\uff09"),
    ("\u767e\u5b9f", "\u767e\u8ed2\u5e97", "\u30a2\u30f3\u30b0\u30e9", "#333333", "\u3086\u3046\u3059\u3051"),
    ("\u30d5\u30a1\u30a4\u30e4\u30fc", "\u30d5\u30a1\u30a4\u30e4\u30fc\u901a\u308a", "\u4e00\u5339\u72fc\u30fb\u6700\u5f37", "#6B2D8B", "\u3044\u304a\u3046"),
    ("\u30df\u30e4\u30b7\u30bf", "\u5bae\u4e0b\u516c\u5712", "\u30b9\u30b1\u30fc\u30bf\u30fc\u30fb\u81ea\u7531", "#1E6FBF", "\u30ab\u30a4"),
    ("\u30aa\u30af\u30b7\u30d6", "\u5965\u6e0b\u8c37", "\u30a8\u30ea\u30fc\u30c8\u30fb\u7ba1\u7406", "#1A5C30", "\u3072\u304b\u308b"),
    ("\u30b5\u30af\u30e9", "\u685c\u30f6\u4e18", "\u9ad8\u8cb4\u30fb\u5de8\u5927\u8cc7\u672c", "#8B6914", "\u30c4\u30e8\u30b7\uff06\u30d2\u30c8\u30df"),
    ("\u30de\u30eb\u30e4\u30de", "\u5186\u5c71\u753a", "\u30af\u30e9\u30d6\u30fbDJ", "#D4267E", "\u30ca\u30aa"),
    ("\u30c9\u30a6\u30b2\u30f3", "\u9053\u7384\u5742", "\u30ab\u30aa\u30b9\u30fb\u96d1\u591a", "#D4742D", "\u30c6\u30c4"),
    ("\u30b3\u30a6\u30a8\u30f3", "\u516c\u5712\u901a\u308a", "\u30a2\u30fc\u30c8\u30fb\u6f14\u5287", "#888888", "\u30ec\u30f3"),
    ("\u30df\u30e4\u30de\u30b9", "\u5bae\u76ca\u5742", "\u30c6\u30c3\u30af\u30fb\u30c7\u30fc\u30bf", "#C4A800", "\u30b7\u30e5\u30f3"),
]

y -= 3 * mm
row_h = 6 * mm
c.setFillColor(HEADER_BG)
c.rect(MARGIN, y - row_h, CONTENT_W, row_h, fill=1, stroke=0)
col_positions = [MARGIN, MARGIN + 22 * mm, MARGIN + 48 * mm, MARGIN + 80 * mm, MARGIN + 110 * mm]
headers = ["\u6821\u540d", "\u30a8\u30ea\u30a2", "\u30ab\u30eb\u30c1\u30e3\u30fc", "\u30ab\u30e9\u30fc", "\u30ea\u30fc\u30c0\u30fc"]
c.setFillColor(white)
c.setFont(FONT, SMALL_SIZE)
for i, h in enumerate(headers):
    c.drawString(col_positions[i] + 2 * mm, y - row_h + 2 * mm, h)
y -= row_h

for name, area, culture, hex_color, leader in schools:
    accent = HexColor(hex_color)
    c.setFillColor(CARD_BG)
    c.rect(MARGIN, y - row_h, CONTENT_W, row_h, fill=1, stroke=0)
    c.setFillColor(accent)
    c.rect(col_positions[3] + 2 * mm, y - row_h + 1.5 * mm, 3 * mm, 3 * mm, fill=1, stroke=0)
    c.setFont(FONT, BODY_SIZE)
    c.drawString(col_positions[0] + 2 * mm, y - row_h + 2 * mm, name)
    c.setFillColor(black)
    c.setFont(FONT, SMALL_SIZE)
    c.drawString(col_positions[1] + 2 * mm, y - row_h + 2 * mm, area)
    c.drawString(col_positions[2] + 2 * mm, y - row_h + 2 * mm, culture)
    c.drawString(col_positions[3] + 7 * mm, y - row_h + 2 * mm, hex_color)
    c.drawString(col_positions[4] + 2 * mm, y - row_h + 2 * mm, leader)
    y -= row_h


# ═══════════════════════════════════════════════════════
# PAGE 4: FIRST SEASON STORY
# ═══════════════════════════════════════════════════════
new_page()
y = PAGE_H - MARGIN
c.setFillColor(RED)
c.rect(MARGIN - 3 * mm, MARGIN, 2 * mm, PAGE_H - 2 * MARGIN, fill=1, stroke=0)

y = draw_header_bar(y, "FIRST SEASON STORY", color=RED)
y -= 5 * mm
c.setFillColor(black)
c.setFont(FONT, 11)
c.drawString(MARGIN, y, "\u30d5\u30a1\u30fc\u30b9\u30c8\u30b7\u30fc\u30ba\u30f3\u516820\u8a71\u300c\u59cb\u307e\u308a\u306e\u708e\u300d")
y -= 6 * mm
c.setFillColor(GRAY)
c.setFont(FONT, BODY_SIZE)
c.drawString(MARGIN, y, "\u5e8f\u7ae0 \u2192 \u30ea\u30f3\u306e\u8ee2\u6821 \u2192 \u6e0b\u4e00vs\u767e\u5b9f\u306e\u5bfe\u7acb\u6fc0\u5316 \u2192 \u5185\u90e8\u5d29\u58ca\u3068\u518d\u8d77 \u2192 \u6700\u7d42\u6c7a\u6226\u3068\u65b0\u305f\u306a\u8105\u5a01")

acts = [
    ("\u7b2c\u4e00\u5e55\uff081\u301c5\u8a71\uff09\u5747\u8861\u3068\u51fa\u4f1a\u3044",
     "\u308b\u3044\u3068\u30ea\u30f3\u306e\u51fa\u4f1a\u3044\u3002\u767e\u5b9f\u306e\u3086\u3046\u3059\u3051\u304c\u30ea\u30f3\u306e\u596a\u9084\u3092\u5ba3\u8a00\u3057\u3001\u51b7\u6226\u304c\u5d29\u308c\u308b\u3002", HexColor('#1E6FBF')),
    ("\u7b2c\u4e8c\u5e55\uff086\u301c10\u8a71\uff09\u6fc0\u7a81\u3068\u62e1\u5927",
     "\u30ad\u30e3\u30c3\u30c8\u30b9\u30c8\u30ea\u30fc\u30c8\u3067\u6e0b\u4e00vs\u767e\u5b9f\u304c\u6fc0\u7a81\u3002\u3044\u304a\u3046\u53c2\u6226\u3002\u3086\u3046\u305f\u306e\u88cf\u5207\u308a\u306e\u7a2e\u304c\u82bd\u751f\u3048\u308b\u3002", GOLD),
    ("\u7b2c\u4e09\u5e55\uff0811\u301c15\u8a71\uff09\u5d29\u58ca\u3068\u7d76\u671b",
     "\u3086\u3046\u305f\u306e\u88cf\u5207\u308a\u3067\u6e0b\u4e00\u5d29\u58ca\u3002Mr.K\u3068\u306e\u9082\u9005\u3067\u308b\u3044\u304c\u518d\u8d77\u3002O.K.thanks\u65b0\u66f2\u304c\u6d41\u308c\u308b\u3002", RED),
    ("\u7b2c\u56db\u5e55\uff0816\u301c20\u8a71\uff09\u6700\u7d42\u6c7a\u6226\u3068\u672a\u6765",
     "\u5168\u9762\u6226\u4e89\u3002\u308b\u3044vs\u3086\u3046\u3059\u3051\u6700\u7d42\u6c7a\u6226\u3067\u52dd\u5229\u3002\u30b5\u30af\u30e9\u304c\u6b21\u306a\u308b\u8105\u5a01\u3068\u3057\u3066\u767b\u5834\u3002", HexColor('#1A5C30')),
]

y -= 8 * mm
for title, desc, accent in acts:
    card_h = 18 * mm
    c.setFillColor(CARD_BG)
    c.rect(MARGIN, y - card_h, CONTENT_W, card_h, fill=1, stroke=0)
    c.setFillColor(accent)
    c.rect(MARGIN, y, CONTENT_W, 0.8 * mm, fill=1, stroke=0)
    c.setFont(FONT, 9)
    c.drawString(MARGIN + 3 * mm, y - 5 * mm, title)
    c.setFillColor(GRAY)
    c.setFont(FONT, BODY_SIZE)
    c.drawString(MARGIN + 3 * mm, y - 12 * mm, desc)
    y -= card_h + 4 * mm


# ═══════════════════════════════════════════════════════
# PAGE 5: CAST
# ═══════════════════════════════════════════════════════
new_page()
y = PAGE_H - MARGIN
c.setFillColor(RED)
c.rect(MARGIN - 3 * mm, MARGIN, 2 * mm, PAGE_H - 2 * MARGIN, fill=1, stroke=0)

y = draw_header_bar(y, "CAST \u2500 \u30ad\u30e3\u30b9\u30c8", color=RED)

cast_data = [
    ("\u308b\u3044", "\u6e0b\u4e00", "\u4e3b\u4eba\u516c\u30fb\u4e0d\u5c48\u306e\u7cbe\u795e", "#CC0000"),
    ("\u306f\u308b", "\u6e0b\u4e00", "\u76f8\u68d2\u30fb\u7279\u653b\u968a\u9577", "#CC0000"),
    ("\u308a\u3087\u3046\u3059\u3051", "\u6e0b\u4e00", "\u5be1\u9ed9\u306a\u6700\u5f37\u30a4\u30b1\u30e1\u30f3", "#CC0000"),
    ("\u3086\u3046\u305f", "\u6e0b\u4e00", "\u53c2\u8b00\u2192\u88cf\u5207\u308a", "#CC0000"),
    ("\u3057\u304a\u3093", "\u6e0b\u4e00", "\u5144\u8cb4\u808c\u2192\u3086\u3046\u305f\u306b\u540c\u8abf", "#CC0000"),
    ("\u3086\u3046\u3059\u3051", "\u767e\u5b9f", "\u30ea\u30fc\u30c0\u30fc\u30fb\u30ab\u30ea\u30b9\u30de", "#333333"),
    ("\u3068\u3046\u307e", "\u767e\u5b9f", "\u526f\u30ea\u30fc\u30c0\u30fc", "#333333"),
    ("\u306f\u308c", "\u767e\u5b9f", "\u5207\u308a\u8fbc\u307f\u968a\u9577", "#333333"),
    ("\u3053\u3046\u306e\u3059\u3051", "\u767e\u5b9f", "\u60c5\u5831\u5c4b", "#333333"),
    ("\u305b\u308a\u306a", "\u767e\u5b9f", "\u30ae\u30e3\u30eb\u30fb\u7d05\u4e00\u70b9", "#333333"),
    ("\u3044\u304a\u3046", "\u30d5\u30a1\u30a4\u30e4\u30fc", "\u4e00\u5339\u72fc\u30fb\u6e0b\u8c37\u6700\u5f37", "#6B2D8B"),
    ("\u3072\u304b\u308b", "\u30aa\u30af\u30b7\u30d6", "\u7ba1\u7406\u8005\u30ea\u30fc\u30c0\u30fc", "#1A5C30"),
]

y -= 3 * mm
row_h = 6 * mm
c.setFillColor(HEADER_BG)
c.rect(MARGIN, y - row_h, CONTENT_W, row_h, fill=1, stroke=0)
cast_cols = [MARGIN, MARGIN + 25 * mm, MARGIN + 55 * mm]
cast_headers = ["\u540d\u524d", "\u6240\u5c5e", "\u5f79\u5272"]
c.setFillColor(white)
c.setFont(FONT, SMALL_SIZE)
for i, h in enumerate(cast_headers):
    c.drawString(cast_cols[i] + 2 * mm, y - row_h + 2 * mm, h)
y -= row_h

for name, school, role, hex_color in cast_data:
    accent = HexColor(hex_color)
    c.setFillColor(CARD_BG)
    c.rect(MARGIN, y - row_h, CONTENT_W, row_h, fill=1, stroke=0)
    c.setFillColor(accent)
    c.rect(MARGIN, y - row_h, 1.5 * mm, row_h, fill=1, stroke=0)
    c.setFont(FONT, BODY_SIZE)
    c.drawString(cast_cols[0] + 3 * mm, y - row_h + 2 * mm, name)
    c.setFillColor(GRAY)
    c.setFont(FONT, SMALL_SIZE)
    c.drawString(cast_cols[1] + 2 * mm, y - row_h + 2 * mm, school)
    c.setFillColor(black)
    c.drawString(cast_cols[2] + 2 * mm, y - row_h + 2 * mm, role)
    y -= row_h

y -= 5 * mm
c.setFillColor(GRAY)
c.setFont(FONT, SMALL_SIZE)
c.drawString(MARGIN, y, "Female cast (tentative):  \u308a\u3042\u3001\u306f\u308b\u306a\u3001\u307e\u3046\u3001\u306f\u308b\u3042\u3001\u3058\u3085\u306a\u3001\u306a\u308b\u305b")


# ═══════════════════════════════════════════════════════
# PAGE 6: O.K.thanks
# ═══════════════════════════════════════════════════════
new_page()
y = PAGE_H - MARGIN
c.setFillColor(GOLD)
c.rect(MARGIN - 3 * mm, MARGIN, 2 * mm, PAGE_H - 2 * MARGIN, fill=1, stroke=0)

y = draw_header_bar(y, "O.K.thanks \u2500 \u4f1d\u8aac\u306e\u5b58\u5728", color=HexColor('#8B6914'))
y -= 5 * mm
y = draw_wrapped_text(MARGIN, y,
    "\u304b\u3064\u3066\u6e0b\u8c37\u3092\u7d71\u4e00\u3057\u305f\u30c1\u30fc\u30e0\u3068\u5bfe\u7b49\u306b\u6e21\u308a\u5408\u3063\u305f\u4f1d\u8aac\u7684\u5b58\u5728\u3002\n\u7269\u8a9e\u306e\u4e2d\u3067\u306f\u300c\u30b9\u30da\u30b7\u30e3\u30eb\u30b5\u30f3\u30af\u30b9\u300d\u7684\u306b\u767b\u5834\u3057\u3001\u4e3b\u4eba\u516c\u308b\u3044\u306e\u5fc3\u306b\u706b\u3092\u706f\u3059\u3002\n\u30ea\u30a2\u30eb\u3067\u306f\u30d7\u30ed\u30b8\u30a7\u30af\u30c8\u306e\u30b9\u30dd\u30f3\u30b5\u30fc\u3068\u3057\u3066\u3001\u65b0\u66f2\u304c\u30d5\u30a1\u30fc\u30b9\u30c8\u30b7\u30fc\u30ba\u30f3\u306e\u30c6\u30fc\u30de\u66f2\u306b\u3002",
    size=BODY_SIZE, color=black)

y -= 8 * mm
chars = [
    ("Mr.K\uff08\u706b\u3092\u706f\u3059\u8005\uff09",
     "\u7b2c14\u8a71\u3067\u308b\u3044\u304c\u7d76\u671b\u306e\u5e95\u306b\u3044\u308b\u3068\u304d\u3001\u8def\u5730\u88cf\u3067\u51fa\u4f1a\u3046\u3002",
     "\u300c\u6e0b\u8c37\u306f\u306a\u3001\u8ab0\u304b\u304c\u5b88\u308d\u3046\u3068\u3059\u308b\u9650\u308a\u3001\u6b7b\u306a\u306d\u3047\u3088\u300d", RED),
    ("Mr.O\uff08\u98a8\u3092\u9001\u308b\u8005\uff09",
     "\u7b2c15\u8a71\u3067\u308b\u3044\u306e\u518d\u8d77\u3092\u5f8c\u62bc\u3057\u3002",
     "\u300c\u3044\u3044\u9854\u3057\u3066\u3093\u306a\u300d", GOLD),
]
for name, scene, quote, accent in chars:
    card_h = 22 * mm
    c.setFillColor(CARD_BG)
    c.rect(MARGIN, y - card_h, CONTENT_W, card_h, fill=1, stroke=0)
    c.setFillColor(accent)
    c.rect(MARGIN, y, CONTENT_W, 0.8 * mm, fill=1, stroke=0)
    c.setFont(FONT, 9)
    c.drawString(MARGIN + 3 * mm, y - 5 * mm, name)
    c.setFillColor(GRAY)
    c.setFont(FONT, BODY_SIZE)
    c.drawString(MARGIN + 3 * mm, y - 11 * mm, scene)
    c.setFillColor(black)
    c.setFont(FONT, 9)
    c.drawString(MARGIN + 3 * mm, y - 17 * mm, quote)
    y -= card_h + 5 * mm


# ═══════════════════════════════════════════════════════
# PAGE 7: ROADMAP
# ═══════════════════════════════════════════════════════
new_page()
y = PAGE_H - MARGIN
c.setFillColor(RED)
c.rect(MARGIN - 3 * mm, MARGIN, 2 * mm, PAGE_H - 2 * MARGIN, fill=1, stroke=0)

y = draw_header_bar(y, "ROADMAP \u2500 3\u30d5\u30a7\u30fc\u30ba", color=RED)

phases = [
    ("PHASE 1 (2026)", "SNS\u59cb\u52d5\u30fb\u8a8d\u77e5\u62e1\u5927",
     "\u300c\u5b66\u5712\u300d\u30b7\u30e7\u30fc\u30c8\u30e0\u30fc\u30d3\u30fc\u3092\u8ef8\u306b\u30ad\u30e3\u30e9\u30af\u30bf\u30fc\u5b9a\u7740\u3002TGC TEEN\u3067\u30c7\u30d3\u30e5\u30fc\u3002\nSNS\u30aa\u30fc\u30ac\u30cb\u30c3\u30af\u62e1\u6563\u3002\u30d5\u30a1\u30fc\u30b9\u30c8\u30b7\u30fc\u30ba\u30f3\u516820\u8a71\u5c55\u958b\u3002", HexColor('#1E6FBF')),
    ("PHASE 2 (2026-27)", "\u77ed\u7de8\u6620\u753b\u30fb\u30e1\u30c7\u30a3\u30a2\u9023\u643a",
     "\u6e0b\u8c37\u821e\u53f0\u306e\u77ed\u7de8\u6620\u753b\u3002\u5730\u4e0a\u6ce2\u9023\u52d5\u4f01\u753b\u3002\n\u30df\u30b9\u30b3\u30f3\u76f8\u4e92\u9001\u5ba2\u3002\u30d5\u30a1\u30f3\u30af\u30e9\u30d6\u8a2d\u7acb\u3002", GOLD),
    ("PHASE 3 (2027-28)", "IP\u78ba\u7acb\u30fbNo.1",
     "\u5730\u4e0a\u6ce2\u30ec\u30ae\u30e5\u30e9\u30fc\u3002\u5927\u624b\u30b5\u30d6\u30b9\u30afPF\u72ec\u5360\u914d\u4fe1\u3002\n\u30e1\u30f3\u30ba\u82b8\u80fd\u90e8\u9580\u30b7\u30a7\u30a21\u4f4d\u3002100\u5104\u5186IP\u4fa1\u5024\u3002", RED),
]

y -= 5 * mm
for phase_title, subtitle, desc, accent in phases:
    card_h = 28 * mm
    c.setFillColor(CARD_BG)
    c.rect(MARGIN, y - card_h, CONTENT_W, card_h, fill=1, stroke=0)
    c.setFillColor(accent)
    c.rect(MARGIN, y, CONTENT_W, 0.8 * mm, fill=1, stroke=0)
    c.setFont(FONT, 9)
    c.drawString(MARGIN + 3 * mm, y - 5 * mm, phase_title)
    c.setFillColor(black)
    c.setFont(FONT, 10)
    c.drawString(MARGIN + 3 * mm, y - 11 * mm, subtitle)
    c.setFillColor(GRAY)
    c.setFont(FONT, BODY_SIZE)
    dy = y - 17 * mm
    for line in desc.split('\n'):
        c.drawString(MARGIN + 3 * mm, dy, line)
        dy -= 4.5 * mm
    y -= card_h + 5 * mm


# ═══════════════════════════════════════════════════════
# PAGE 8: KPI + CLOSING
# ═══════════════════════════════════════════════════════
new_page()
y = PAGE_H - MARGIN
c.setFillColor(GOLD)
c.rect(MARGIN - 3 * mm, MARGIN, 2 * mm, PAGE_H - 2 * MARGIN, fill=1, stroke=0)

y = draw_header_bar(y, "KPI", color=HexColor('#8B6914'))

kpis = [
    ("\u00a5100\u5104", "IP VALUE TARGET (2028\u5e74)"),
    ("100\u4e07", "YouTube\u767b\u9332\u8005"),
    ("No.1", "\u30e1\u30f3\u30ba\u82b8\u80fd\u90e8\u9580"),
    ("300\u4e07/\u6708", "\u6708\u6b21\u5236\u4f5c\u4e88\u7b97"),
]

y -= 3 * mm
col_w = (CONTENT_W - 6 * mm) / 4
for i, (num, label) in enumerate(kpis):
    bx = MARGIN + i * (col_w + 2 * mm)
    c.setFillColor(CARD_BG)
    c.rect(bx, y - 20 * mm, col_w, 20 * mm, fill=1, stroke=0)
    c.setFillColor(GOLD)
    c.setFont(FONT, 14)
    tw = c.stringWidth(num, FONT, 14)
    c.drawString(bx + (col_w - tw) / 2, y - 9 * mm, num)
    c.setFillColor(black)
    c.setFont(FONT, LABEL_SIZE)
    tw = c.stringWidth(label, FONT, LABEL_SIZE)
    c.drawString(bx + (col_w - tw) / 2, y - 16 * mm, label)

y -= 28 * mm
c.setFillColor(black)
c.setFont(FONT, 9)
c.drawString(MARGIN, y, "Competitive Advantages")
y -= 6 * mm

advantages = [
    ("\u2460  \u6e0b\u8c37\u3068\u3044\u3046\u5727\u5012\u7684\u30d6\u30e9\u30f3\u30c9", RED),
    ("\u2461  Z\u30fb\u03b1\u4e16\u4ee3\u5973\u6027\u306e\u6d88\u8cbb\u884c\u52d5\uff08\u63a8\u3057\u8ab2\u91d1\u6587\u5316\uff09", GOLD),
    ("\u2462  HJ\u306e\u72ec\u81ea\u30cd\u30c3\u30c8\u30ef\u30fc\u30af\uff08\u30df\u30b9\u30b3\u30f3 \u00d7 TGC TEEN \u00d7 \u5730\u4e0a\u6ce2\uff09", HexColor('#1E6FBF')),
]
for text, accent in advantages:
    c.setFillColor(CARD_BG)
    c.rect(MARGIN, y - 7 * mm, CONTENT_W, 7 * mm, fill=1, stroke=0)
    c.setFillColor(accent)
    c.rect(MARGIN, y - 7 * mm, 1.5 * mm, 7 * mm, fill=1, stroke=0)
    c.setFillColor(black)
    c.setFont(FONT, BODY_SIZE)
    c.drawString(MARGIN + 4 * mm, y - 5 * mm, text)
    y -= 9 * mm

y -= 15 * mm
c.setFillColor(RED)
c.rect(MARGIN, y, CONTENT_W, 0.5 * mm, fill=1, stroke=0)
y -= 12 * mm

c.setFillColor(black)
c.setFont(FONT, 16)
closing = "\u6e0b\u4e00\uff08\u30b7\u30d6\u30ef\u30f3\uff09"
tw = c.stringWidth(closing, FONT, 16)
c.drawString((PAGE_W - tw) / 2, y, closing)

y -= 10 * mm
c.setFillColor(GRAY)
c.setFont(FONT, 10)
sub = "\u65e5\u672c\u6700\u5927\u306e\u30e1\u30f3\u30baIP\u5275\u51fa\u88c5\u7f6e\u3078\u3002"
tw = c.stringWidth(sub, FONT, 10)
c.drawString((PAGE_W - tw) / 2, y, sub)

y -= 10 * mm
c.setFillColor(GOLD)
c.setFont(FONT, 8)
sponsor = "\u30c6\u30fc\u30de\u66f2:  O.K.thanks \u65b0\u66f2   |   \u30b9\u30dd\u30f3\u30b5\u30fc:  O.K.thanks"
tw = c.stringWidth(sponsor, FONT, 8)
c.drawString((PAGE_W - tw) / 2, y, sponsor)

y -= 10 * mm
c.setFillColor(RED)
c.rect(PAGE_W / 2 - 30 * mm, y, 60 * mm, 0.5 * mm, fill=1, stroke=0)

y -= 10 * mm
c.setFillColor(GRAY)
c.setFont(FONT, 7)
conf = "CONFIDENTIAL  \u00a9  HJ Inc. All Rights Reserved."
tw = c.stringWidth(conf, FONT, 7)
c.drawString((PAGE_W - tw) / 2, y, conf)

c.save()
print(f"Saved: {OUT}")
