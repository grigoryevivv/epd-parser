"""Генерирует icon.ico (запускается один раз, результат лежит в репозитории)."""
from PIL import Image, ImageDraw

S = 256
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)
d.rounded_rectangle((8, 8, S - 8, S - 8), radius=52, fill=(47, 125, 246))
# лист
d.rounded_rectangle((52, 40, 168, 196), radius=10, fill=(255, 255, 255))
d.polygon([(136, 40), (168, 72), (136, 72)], fill=(200, 215, 240))
for y in (100, 124, 148):
    d.line((72, y, 148, y), fill=(47, 125, 246), width=7)
# чертёж-уголок
d.rounded_rectangle((120, 128, 220, 220), radius=14, fill=(229, 72, 77))
d.line((140, 196, 200, 196), fill="white", width=8)
d.line((140, 196, 140, 148), fill="white", width=8)
d.ellipse((160, 150, 196, 186), outline="white", width=7)
img.save("icon.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
img.save("icon.png")
