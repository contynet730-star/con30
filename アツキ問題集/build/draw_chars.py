"""キャラクターイラストを描く（Pillow）。4倍で描いて縮小し、なめらかにする。"""
import math, sys
from PIL import Image, ImageDraw

OUT = sys.argv[1]
S = 4  # supersampling

def canvas(w=400, h=400):
    im = Image.new("RGBA", (w * S, h * S), (0, 0, 0, 0))
    return im, ImageDraw.Draw(im)

def save(im, name, w=400, h=400):
    im.resize((w, h), Image.LANCZOS).save(f"{OUT}/{name}.png")

def P(*v):  # scale coords
    return [x * S for x in v]

def ell(d, cx, cy, rx, ry, fill, outline=None, width=0):
    d.ellipse(P(cx - rx, cy - ry, cx + rx, cy + ry), fill=fill,
              outline=outline, width=width * S if outline else 0)

def eye(d, cx, cy, r=22, look=(3, 4), white=(255, 255, 255)):
    ell(d, cx, cy, r, r * 1.1, white, (40, 25, 15), 4)
    ell(d, cx + look[0], cy + look[1], r * 0.5, r * 0.55, (30, 20, 15))
    ell(d, cx + look[0] - r * 0.18, cy + look[1] - r * 0.22, r * 0.17, r * 0.17, (255, 255, 255))

def smile(d, cx, cy, w, h, color=(70, 30, 20), width=7, open_mouth=True):
    if open_mouth:
        d.chord(P(cx - w, cy - h, cx + w, cy + h), 0, 180, fill=(170, 40, 50),
                outline=color, width=width * S)
        d.chord(P(cx - w * 0.5, cy + h * 0.2, cx + w * 0.5, cy + h * 1.0), 180, 360,
                fill=(240, 120, 130))
    else:
        d.arc(P(cx - w, cy - h, cx + w, cy + h), 20, 160, fill=color, width=width * S)

def cheeks(d, xs, cy, r=16):
    for x in xs:
        ell(d, x, cy, r * 1.3, r * 0.8, (255, 150, 150, 160))

BROWN = (150, 95, 50)
BROWN_D = (85, 50, 25)

def poop_body(d, cx=200, base=360, scale=1.0, fill=BROWN, line=BROWN_D):
    s = scale
    tiers = [  # (cy, rx, ry)
        (base - 55 * s, 165 * s, 62 * s),
        (base - 140 * s, 125 * s, 55 * s),
        (base - 215 * s, 85 * s, 45 * s),
    ]
    for cy, rx, ry in tiers:
        ell(d, cx, cy, rx, ry, fill, line, 7)
        ell(d, cx - rx * 0.45, cy - ry * 0.35, rx * 0.28, ry * 0.22, (255, 255, 255, 70))
    # tip
    tip = [(cx - 40 * s, base - 245 * s), (cx + 30 * s, base - 245 * s),
           (cx + 22 * s, base - 290 * s), (cx + 45 * s, base - 320 * s),
           (cx - 5 * s, base - 300 * s)]
    d.polygon([(x * S, y * S) for x, y in tip], fill=fill, outline=line, width=7 * S)
    ell(d, cx, base - 245 * s, 52 * s, 28 * s, fill)
    return tiers

# 1) ウンチー博士 --------------------------------------------------------
im, d = canvas()
tiers = poop_body(d)
cy = tiers[1][0]
# 丸メガネ
for x in (160, 240):
    eye(d, x, cy - 5, 22)
    ell(d, x, cy - 5, 36, 36, None, (40, 40, 40), 6)
d.line(P(196, cy - 8, 204, cy - 8), fill=(40, 40, 40), width=6 * S)
# 白いヒゲ
for sgn in (-1, 1):
    pts = [(200, cy + 40), (200 + sgn * 30, cy + 30), (200 + sgn * 75, cy + 45),
           (200 + sgn * 60, cy + 58), (200 + sgn * 25, cy + 55)]
    d.polygon([(x * S, y * S) for x, y in pts], fill=(250, 250, 245), outline=(120, 120, 120), width=3 * S)
smile(d, 200, tiers[0][0] - 10, 30, 22)
cheeks(d, (130, 270), tiers[0][0] - 25)
# 博士の帽子
d.polygon(P(130, 75, 270, 75, 200, 45), fill=(30, 30, 60))
d.rectangle(P(165, 75, 235, 95), fill=(30, 30, 60))
d.line(P(255, 72, 270, 120), fill=(240, 200, 40), width=5 * S)
ell(d, 270, 125, 8, 8, (240, 200, 40))
save(im, "hakase")

# 2) 小さいうんこアイコン（サッカーボールつき）-----------------------------
im, d = canvas()
tiers = poop_body(d, base=370)
cy = tiers[1][0]
eye(d, 165, cy - 5, 20, (4, 2)); eye(d, 235, cy - 5, 20, (4, 2))
smile(d, 200, tiers[0][0] - 15, 34, 26)
cheeks(d, (125, 275), tiers[0][0] - 25)
save(im, "unko")

# 3) 雲キャラ共通 ---------------------------------------------------------
def cloud(d, circles, fill, line, lw=8):
    for (x, y, r) in circles:
        ell(d, x, y, r + lw, r + lw, line)
    for (x, y, r) in circles:
        ell(d, x, y, r, r, fill)

CL = [(200, 210, 110), (110, 240, 75), (290, 240, 75), (145, 150, 70),
      (255, 150, 70), (200, 280, 85), (130, 290, 55), (270, 290, 55)]

# ブー太郎（黄緑のおならぐも）
im, d = canvas()
for i, (x, y) in enumerate([(60, 340), (40, 300), (75, 375)]):
    d.line(P(x - 30, y, x - 5, y), fill=(140, 170, 60), width=8 * S)
cloud(d, CL, (205, 225, 110), (110, 135, 40))
eye(d, 165, 205, 22, (5, 0)); eye(d, 240, 205, 22, (5, 0))
d.line(P(140, 165, 185, 175), fill=(60, 70, 20), width=8 * S)  # まゆ
d.line(P(260, 165, 218, 175), fill=(60, 70, 20), width=8 * S)
smile(d, 202, 265, 40, 30)
# 背番号ハチマキ
d.rectangle(P(105, 120, 295, 140), fill=(220, 40, 40))
save(im, "butaro")

# スカシ丸（半とうめいの忍者ぐも）
im, d = canvas()
cloud(d, CL, (215, 220, 230, 170), (130, 140, 160, 220))
d.rectangle(P(95, 175, 305, 235), fill=(35, 35, 50, 235))   # 忍者マスク
d.polygon(P(300, 190, 360, 160, 350, 205), fill=(35, 35, 50, 235))
eye(d, 165, 205, 18, (6, 0)); eye(d, 240, 205, 18, (6, 0))
d.arc(P(175, 255, 230, 285), 20, 160, fill=(80, 80, 100), width=6 * S)
save(im, "sukashi")

# ヘッピー（オレンジのぷっくりぐも＋音ぷ）
im, d = canvas()
cloud(d, CL, (255, 190, 110), (190, 110, 40))
eye(d, 165, 210, 20, (0, 6)); eye(d, 238, 210, 20, (0, 6))
d.arc(P(175, 250, 230, 290), 200, 340, fill=(120, 60, 20), width=7 * S)  # きんちょう口
ell(d, 125, 255, 22, 13, (255, 130, 130, 180)); ell(d, 280, 255, 22, 13, (255, 130, 130, 180))
for (x, y) in [(330, 90), (360, 140)]:
    ell(d, x, y, 14, 11, (60, 60, 60))
    d.line(P(x + 12, y, x + 12, y - 50), fill=(60, 60, 60), width=5 * S)
    d.line(P(x + 12, y - 50, x + 30, y - 38), fill=(60, 60, 60), width=5 * S)
# あせ
d.polygon(P(300, 160, 290, 185, 310, 185), fill=(120, 190, 255))
ell(d, 300, 188, 10, 10, (120, 190, 255))
save(im, "heppy")

# 4) プゥ子（もも型GK）----------------------------------------------------
im, d = canvas()
ell(d, 200, 225, 140, 135, (255, 175, 185), (200, 90, 110), 7)
ell(d, 150, 190, 55, 45, (255, 220, 225, 150))
d.arc(P(150, 70, 300, 240), 170, 250, fill=(200, 90, 110), width=6 * S)  # ももの線
# 葉っぱ
d.polygon(P(200, 92, 150, 50, 120, 70, 165, 100), fill=(90, 170, 80), outline=(50, 110, 40), width=4 * S)
d.polygon(P(205, 92, 255, 45, 285, 70, 240, 100), fill=(90, 170, 80), outline=(50, 110, 40), width=4 * S)
eye(d, 160, 220, 22, (2, 3)); eye(d, 240, 220, 22, (2, 3))
for x in (140, 148, 156):  # まつげ
    d.line(P(x, 194, x - 6, 182), fill=(40, 25, 15), width=4 * S)
for x in (244, 252, 260):
    d.line(P(x, 194, x + 6, 182), fill=(40, 25, 15), width=4 * S)
smile(d, 200, 285, 28, 20)
for x in (45, 355):  # GKグローブ
    ell(d, x, 260, 42, 48, (255, 140, 30), (170, 80, 10), 6)
    d.line(P(x - 20, 240, x + 20, 240), fill=(170, 80, 10), width=5 * S)
save(im, "puko")

# 5) 大魔王ベンピー（カチカチ四角うんこ）----------------------------------
im, d = canvas()
DK, DL = (95, 70, 55), (45, 30, 22)
blocks = [(55, 260, 345, 370), (90, 165, 310, 265), (125, 85, 275, 170)]
for (x1, y1, x2, y2) in blocks:
    d.rounded_rectangle(P(x1, y1, x2, y2), radius=14 * S, fill=DK, outline=DL, width=8 * S)
    d.line(P(x1 + 30, y1 + 22, x1 + 70, y1 + 22), fill=(160, 140, 120), width=6 * S)  # ひび
# つの
d.polygon(P(130, 95, 110, 25, 165, 85), fill=(230, 230, 220), outline=DL, width=5 * S)
d.polygon(P(270, 95, 290, 25, 235, 85), fill=(230, 230, 220), outline=DL, width=5 * S)
# おこった目
for (x, sgn) in ((160, 1), (240, -1)):
    ell(d, x, 215, 22, 18, (255, 230, 60), DL, 4)
    ell(d, x, 218, 9, 12, (200, 20, 20))
    d.line(P(x - 30 * sgn, 182, x + 25 * sgn, 198), fill=DL, width=10 * S)
# 口（ギザギザ）
d.rectangle(P(130, 300, 270, 340), fill=(60, 20, 20), outline=DL, width=5 * S)
for i in range(7):
    x = 130 + i * 20
    d.polygon(P(x, 300, x + 20, 300, x + 10, 318), fill=(250, 250, 250))
# マント
save(im, "benpi")

# 6) サッカーボール -------------------------------------------------------
im, d = canvas()
cx, cy, R = 200, 200, 170
ell(d, cx, cy, R, R, (255, 255, 255), (30, 30, 30), 8)
def pent(cx, cy, r, rot):
    return [(cx + r * math.cos(rot + i * 2 * math.pi / 5), cy + r * math.sin(rot + i * 2 * math.pi / 5)) for i in range(5)]
c = pent(cx, cy, 55, -math.pi / 2)
d.polygon([(x * S, y * S) for x, y in c], fill=(30, 30, 30))
mask = Image.new("L", im.size, 0)
ImageDraw.Draw(mask).ellipse(P(cx - R + 6, cy - R + 6, cx + R - 6, cy + R - 6), fill=255)
layer = Image.new("RGBA", im.size, (0, 0, 0, 0))
ld = ImageDraw.Draw(layer)
for i in range(5):
    a = -math.pi / 2 + i * 2 * math.pi / 5
    vx, vy = c[i]
    ox, oy = cx + 150 * math.cos(a), cy + 150 * math.sin(a)
    ld.line([(vx * S, vy * S), (ox * S, oy * S)], fill=(30, 30, 30), width=7 * S)
    p = pent(cx + 175 * math.cos(a), cy + 175 * math.sin(a), 50, a + math.pi)
    ld.polygon([(x * S, y * S) for x, y in p], fill=(30, 30, 30))
    b = a + math.pi / 5
    ld.line([(ox * S, oy * S), ((cx + 165 * math.cos(b)) * S, (cy + 165 * math.sin(b)) * S)], fill=(30, 30, 30), width=7 * S)
    b2 = a - math.pi / 5
    ld.line([(ox * S, oy * S), ((cx + 165 * math.cos(b2)) * S, (cy + 165 * math.sin(b2)) * S)], fill=(30, 30, 30), width=7 * S)
im.paste(layer, (0, 0), Image.composite(layer, Image.new("RGBA", im.size), mask).split()[3])
d = ImageDraw.Draw(im)
ell(d, cx, cy, R, R, None, (30, 30, 30), 8)
save(im, "ball")

# 7) トロフィー（表彰状用）-----------------------------------------------
im, d = canvas()
G, GD = (250, 200, 50), (170, 120, 10)
d.chord(P(100, 40, 300, 260), 0, 180, fill=G, outline=GD, width=7 * S)
d.rectangle(P(100, 40, 300, 150), fill=G, outline=GD, width=7 * S)
d.rectangle(P(107, 47, 293, 155), fill=G)
for sgn in (-1, 1):
    x = 200 + sgn * 120
    d.arc(P(x - 45, 60, x + 45, 160), 270 if sgn > 0 else 90, 90 if sgn > 0 else 270, fill=GD, width=12 * S)
d.rectangle(P(180, 255, 220, 310), fill=G, outline=GD, width=6 * S)
d.rounded_rectangle(P(120, 305, 280, 360), radius=10 * S, fill=(140, 90, 50), outline=(90, 55, 25), width=6 * S)
# トロフィーの上にうんこ
sub = Image.open(f"{OUT}/unko.png").resize((150 * S, 150 * S), Image.LANCZOS)
im.alpha_composite(sub, (125 * S, 60 * S))
save(im, "trophy")
print("ok")
