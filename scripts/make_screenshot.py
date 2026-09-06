from PIL import Image, ImageDraw, ImageFont
import os

# Canvas size matching a typical Excel task pane (narrow, tall)
W, H = 380, 700
BG = "#12161c"
INK2 = "#1c232d"
PAPER = "#f3ede3"
COPPER = "#c4845a"
LEDGER = "#3d7a5a"
MUTE = "#8b8478"
DARK_BORDER = "#3a332c"

img = Image.new("RGB", (W, H), BG)
draw = ImageDraw.Draw(img)

def get_font(size, bold=False):
    """Try a few font options; fall back to default."""
    names = ["SF Pro Text", "Helvetica Neue", "Arial", "DejaVuSans"]
    for name in names:
        try:
            return ImageFont.truetype(name, size)
        except:
            continue
    return ImageFont.load_default()

font_small = get_font(11)
font_body = get_font(13)
font_header = get_font(22)
font_mono = get_font(11)
font_chip = get_font(11)

# ── Header ──
draw.text((16, 14), "EXCEL ANALYST", font=font_small, fill=COPPER)
draw.text((16, 30), "Crunched", font=font_header, fill=PAPER)
draw.text((16, 62), "Working conversation", font=font_small, fill=MUTE)
draw.rectangle([(16, 84), (W-16, 86)], fill=COPPER)

y = 100

def user_bubble(text, y):
    tw, th = draw.textbbox((0, 0), text, font=font_body)[2:]
    pad = 10
    bw, bh = tw + pad*2, th + pad*2
    x = W - 16 - bw
    draw.rounded_rectangle([(x, y), (x+bw, y+bh)], radius=10, fill=PAPER)
    draw.text((x+pad, y+pad), text, font=font_body, fill=BG)
    return y + bh + 10

def assistant_bubble(text, y):
    lines = []
    words = text.split()
    line = ""
    max_w = W - 56
    for w in words:
        test = line + " " + w if line else w
        tw, _ = draw.textbbox((0, 0), test, font=font_body)[2:]
        if tw > max_w:
            lines.append(line)
            line = w
        else:
            line = test
    lines.append(line)
    
    pad = 10
    line_h = draw.textbbox((0, 0), "Ag", font=font_body)[3]
    bh = len(lines) * line_h + pad*2
    bw = W - 48
    draw.rounded_rectangle([(16, y), (16+bw, y+bh)], radius=10, fill=INK2)
    draw.line([(16, y), (16, y+bh)], fill=COPPER, width=2)
    
    for i, line in enumerate(lines):
        draw.text((16+pad, y+pad+i*line_h), line, font=font_body, fill=PAPER)
    return y + bh + 10

def tool_card(name, summary, y):
    pad = 8
    line_h = draw.textbbox((0, 0), "Ag", font=font_mono)[3]
    bh = line_h + pad*2 + 4
    draw.rectangle([(16, y), (W-16, y+bh)], fill=INK2, outline=DARK_BORDER, width=1)
    draw.line([(16, y), (16, y+bh)], fill=LEDGER, width=2)
    draw.text((16+pad, y+pad), name, font=font_mono, fill=LEDGER)
    draw.text((120, y+pad), summary, font=font_mono, fill="#b8b0a4")
    return y + bh + 10

def suggestion_chips(chips, y):
    x = 18
    for chip in chips:
        tw, th = draw.textbbox((0, 0), chip, font=font_chip)[2:]
        pad = 8
        cw, ch = tw + pad*2, th + pad + 4
        draw.rounded_rectangle([(x, y), (x+cw, y+ch)], radius=999, outline=DARK_BORDER, width=1)
        draw.text((x+pad, y+3), chip, font=font_chip, fill=MUTE)
        x += cw + 6
    return y + 28 + 10

# ── Chat content ──
y = user_bubble("How big is this workbook?", y)
y = tool_card("Scanned workbook", "Budget 7×4 · Data 5000×12", y)
y = assistant_bubble("The workbook has 2 sheets. Budget is a small 7-row model. Data is a 5,000-row table with 12 columns.", y)
y = suggestion_chips(["Show formulas", "Find errors", "Add a total"], y)
y = user_bubble("Add a total row to Budget", y)
y = tool_card("Wrote range", "Budget!A7:D7", y)
y = assistant_bubble("Done. Row 7 now shows 'Total' with SUM formulas in B7:D7.", y)
y = suggestion_chips(["Verify the total", "Show the formulas"], y)

# ── Composer area ──
draw.rectangle([(16, H-70), (W-56, H-30)], fill=INK2, outline=DARK_BORDER, width=1)
draw.text((22, H-58), "Ask Crunched...", font=font_small, fill=MUTE)
draw.rounded_rectangle([(W-48, H-70), (W-16, H-30)], radius=6, fill=COPPER)
draw.text((W-40, H-56), "Send", font=font_small, fill=BG)

# ── Selection pill ──
draw.text((16, H-92), "Crunched can see Budget!B2:D5", font=font_mono, fill=MUTE)

output = "/Users/junseki/Documents/GitHub/crunched-kiss/docs/screenshot.png"
img.save(output, "PNG")
print(f"Saved screenshot to {output}")
