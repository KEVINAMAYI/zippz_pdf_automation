import qrcode
img = qrcode.make('https://app.zippz.work/labels')
img.save("test.png")
print(img)

