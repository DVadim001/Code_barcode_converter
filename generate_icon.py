from PIL import Image, ImageDraw

# Создаём пустую картинку
size = (256, 256)
image = Image.new("RGB", size, "white")
draw = ImageDraw.Draw(image)

# Рисуем рамку
draw.rectangle([(0, 0), (255, 255)], outline="black", width=5)

# Рисуем простейший баркод внутри
for i in range(20, 236, 10):
    draw.rectangle([(i, 40), (i + 5, 220)], fill="black")

# Сохраняем как иконку
image.save("icon.ico", format="ICO", sizes=[(256, 256)])

print("Иконка icon.ico успешно создана!")
