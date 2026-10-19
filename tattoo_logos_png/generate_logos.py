from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import math

out_dir = Path(r"c:\Users\Craig\GIT\tattoo_logos_png")
out_dir.mkdir(parents=True, exist_ok=True)

W = H = 1024
BG = (255, 255, 255, 255)

font_candidates = [
    "arial.ttf",
    "segoeui.ttf",
    "calibri.ttf",
    "DejaVuSans.ttf",
]


def get_font(size):
    for f in font_candidates:
        try:
            return ImageFont.truetype(f, size)
        except Exception:
            continue
    return ImageFont.load_default()


def center_text(draw, text, y, size, fill=(0, 0, 0, 255)):
    font = get_font(size)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    draw.text(((W - tw) // 2, y), text, font=font, fill=fill)


def draw_ring_badge(draw, title, subtitle=None):
    cx, cy = W // 2, H // 2
    outer_r = 410
    inner_r = 320
    draw.ellipse(
        (cx - outer_r, cy - outer_r, cx + outer_r, cy + outer_r),
        outline=(0, 0, 0, 255),
        width=8,
        fill=(245, 245, 245, 255),
    )
    draw.ellipse(
        (cx - inner_r, cy - inner_r, cx + inner_r, cy + inner_r),
        outline=(0, 0, 0, 255),
        width=6,
        fill=(255, 255, 255, 255),
    )
    center_text(draw, title, 95, 56)
    if subtitle:
        center_text(draw, subtitle, 860, 42)


def save(img, name):
    img.save(out_dir / name, format="PNG")


# 1) Male symbol - 1988
img = Image.new("RGBA", (W, H), BG)
d = ImageDraw.Draw(img)
center_text(d, "1988", 120, 84)
cx, cy, r = 460, 560, 180
d.ellipse((cx - r, cy - r, cx + r, cy + r), outline=(20, 90, 180, 255), width=26)
x1, y1 = cx + int(r * 0.7), cy - int(r * 0.7)
x2, y2 = x1 + 165, y1 - 165
d.line((x1, y1, x2, y2), fill=(20, 90, 180, 255), width=26)
d.polygon([(x2, y2), (x2 - 90, y2 + 18), (x2 - 18, y2 + 90)], fill=(20, 90, 180, 255))
center_text(d, "Birth of Son", 860, 46)
save(img, "01_1988_birth_of_son_male_symbol.png")

# 2) Trans Pennine Trail concept
img = Image.new("RGBA", (W, H), BG)
d = ImageDraw.Draw(img)
draw_ring_badge(d, "TRANS PENNINE TRAIL", "Cycling Completion - 2005")
d.arc((170, 280, 860, 760), start=190, end=350, fill=(120, 120, 120, 255), width=16)
d.arc((210, 320, 900, 800), start=190, end=350, fill=(160, 160, 160, 255), width=10)
leaf = [(515, 530), (645, 480), (760, 520), (640, 560)]
d.polygon(leaf, fill=(90, 150, 70, 255), outline=(50, 110, 50, 255))
save(img, "02_2005_trans_pennine_trail_cycle.png")

# 3) Female symbol - July 2006
img = Image.new("RGBA", (W, H), BG)
d = ImageDraw.Draw(img)
center_text(d, "JULY 2006", 120, 80)
cx, cy, r = 512, 500, 170
d.ellipse((cx - r, cy - r, cx + r, cy + r), outline=(210, 80, 140, 255), width=24)
d.line((cx, cy + r, cx, cy + r + 170), fill=(210, 80, 140, 255), width=24)
d.line((cx - 90, cy + r + 110, cx + 90, cy + r + 110), fill=(210, 80, 140, 255), width=24)
center_text(d, "Birth of Daughter", 860, 46)
save(img, "03_2006_birth_of_daughter_female_symbol.png")

# 4) W2W concept
img = Image.new("RGBA", (W, H), BG)
d = ImageDraw.Draw(img)
draw_ring_badge(d, "W2W", "Walney to Wear Cycle Route - Aug 2006")
center_text(d, "Valley to Wear", 450, 42, fill=(50, 50, 50, 255))
pts = [(250, 650), (420, 700), (640, 670), (790, 560)]
d.line(pts, fill=(70, 150, 70, 255), width=24, joint="curve")
d.polygon([(790, 560), (730, 560), (760, 620)], fill=(70, 150, 70, 255))
center_text(d, "W2W", 350, 130, fill=(80, 130, 70, 255))
save(img, "04_2006_w2w_walney_to_wear_cycle.png")

# 5) Coast and Castles concept
img = Image.new("RGBA", (W, H), BG)
d = ImageDraw.Draw(img)
draw_ring_badge(d, "COAST & CASTLES", "Cycling Completion - 2007")
d.polygon([(250, 650), (780, 650), (780, 760), (250, 760)], fill=(70, 120, 190, 255))
d.polygon([(250, 650), (500, 520), (780, 650)], fill=(110, 170, 90, 255))
d.rectangle((560, 470, 710, 620), fill=(180, 180, 180, 255), outline=(50, 50, 50, 255), width=4)
d.rectangle((590, 430, 625, 470), fill=(180, 180, 180, 255), outline=(50, 50, 50, 255), width=4)
d.rectangle((650, 430, 685, 470), fill=(180, 180, 180, 255), outline=(50, 50, 50, 255), width=4)
save(img, "05_2007_coast_and_castles_cycle.png")

# 6) Hadrian's Cycleway concept
img = Image.new("RGBA", (W, H), BG)
d = ImageDraw.Draw(img)
draw_ring_badge(d, "HADRIAN'S CYCLEWAY", "Cycling Completion - 2008")
for i in range(8):
    x = 230 + i * 70
    y = 640 - (i % 2) * 10
    d.rectangle((x, y, x + 64, y + 50), fill=(170, 140, 95, 255), outline=(90, 70, 45, 255), width=3)
d.ellipse((330, 470, 410, 550), outline=(30, 30, 30, 255), width=8)
d.ellipse((520, 470, 600, 550), outline=(30, 30, 30, 255), width=8)
d.line((370, 510, 470, 470), fill=(30, 30, 30, 255), width=8)
d.line((470, 470, 560, 510), fill=(30, 30, 30, 255), width=8)
d.line((470, 470, 450, 540), fill=(30, 30, 30, 255), width=8)
save(img, "06_2008_hadrians_cycleway_cycle.png")

# 7) Cumbria Way concept
img = Image.new("RGBA", (W, H), BG)
d = ImageDraw.Draw(img)
draw_ring_badge(d, "CUMBRIA WAY", "On Foot - 2010")
d.polygon([(250, 680), (470, 430), (670, 680)], fill=(110, 160, 90, 255), outline=(50, 90, 50, 255), width=4)
d.line((512, 680, 520, 520), fill=(205, 180, 120, 255), width=18)
for fx, fy in [(720, 650), (770, 700), (290, 720), (340, 770)]:
    d.ellipse((fx, fy, fx + 36, fy + 52), fill=(40, 40, 40, 255))
save(img, "07_2010_cumbria_way_foot.png")

# 8) West Highland Way concept
img = Image.new("RGBA", (W, H), BG)
d = ImageDraw.Draw(img)
draw_ring_badge(d, "WEST HIGHLAND WAY", "On Foot - 2011")
d.polygon([(180, 700), (380, 420), (560, 700)], fill=(100, 145, 110, 255), outline=(40, 80, 50, 255), width=4)
d.polygon([(430, 700), (620, 390), (840, 700)], fill=(80, 125, 95, 255), outline=(35, 70, 45, 255), width=4)
d.rectangle((180, 700, 840, 760), fill=(85, 120, 175, 255))
for fx, fy in [(430, 790), (490, 830), (550, 790), (610, 830)]:
    d.ellipse((fx, fy, fx + 30, fy + 42), fill=(20, 20, 20, 255))
save(img, "08_2011_west_highland_way_foot.png")

# 9) National Three Peaks concept
img = Image.new("RGBA", (W, H), BG)
d = ImageDraw.Draw(img)
draw_ring_badge(d, "NATIONAL THREE PEAKS", "24 HOUR CHALLENGE - 2018")
d.polygon([(180, 730), (330, 520), (470, 730)], fill=(120, 140, 160, 255), outline=(40, 50, 60, 255), width=4)
d.polygon([(330, 730), (520, 430), (700, 730)], fill=(100, 125, 150, 255), outline=(35, 45, 55, 255), width=4)
d.polygon([(560, 730), (730, 500), (860, 730)], fill=(125, 145, 165, 255), outline=(45, 55, 65, 255), width=4)
route = [(230, 690), (330, 560), (520, 470), (730, 540), (820, 690)]
d.line(route, fill=(70, 150, 80, 255), width=14, joint="curve")
for sx, sy, n in [(330, 560, "1"), (520, 470, "2"), (730, 540, "3")]:
    d.ellipse((sx - 14, sy - 14, sx + 14, sy + 14), fill=(190, 40, 40, 255), outline=(255, 255, 255, 255), width=2)
    d.text((sx - 6, sy - 10), n, font=get_font(20), fill=(255, 255, 255, 255))
for fx, fy in [(250, 790), (305, 835), (740, 790), (795, 835)]:
    d.ellipse((fx, fy, fx + 30, fy + 42), fill=(20, 20, 20, 255))
save(img, "09_2018_national_three_peaks_foot.png")

# Contact sheet
files = sorted([p for p in out_dir.glob("*.png") if p.name != "00_contact_sheet.png"])
thumb_w, thumb_h = 340, 340
cols = 3
rows = math.ceil(len(files) / cols)
sheet = Image.new("RGBA", (cols * thumb_w, rows * thumb_h + 80), (248, 248, 248, 255))
sd = ImageDraw.Draw(sheet)
sd.text((20, 20), "Tattoo Logo Concepts", font=get_font(36), fill=(20, 20, 20, 255))
for idx, p in enumerate(files):
    r, c = divmod(idx, cols)
    x, y = c * thumb_w + 10, r * thumb_h + 70
    im = Image.open(p).convert("RGBA")
    im.thumbnail((thumb_w - 20, thumb_h - 30))
    sheet.alpha_composite(im, (x + (thumb_w - 20 - im.width) // 2, y + (thumb_h - 50 - im.height) // 2))
    label = p.name.split("_", 1)[1].replace(".png", "")[:30]
    sd.text((x + 8, y + thumb_h - 28), label, font=get_font(14), fill=(45, 45, 45, 255))

sheet.save(out_dir / "00_contact_sheet.png", format="PNG")
print(f"Created {len(files) + 1} PNG files in: {out_dir}")
