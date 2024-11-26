import qrcode
img = qrcode.make('hola')
type(img)  # qrcode.image.pil.PilImage
img.save("olo.png")