from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

page_width = 279.4 * mm
page_height = 215.9 * mm
card_width = 63 * mm
card_height = 88 * mm

pdf = canvas.Canvas(
    "horizontal_cardBack_3x3.pdf",
    pagesize=(page_width, page_height),
)
pdf.drawImage(
    ImageReader("card-back.jpg"),
    (page_width - card_width) / 2,
    (page_height - card_height) / 2,
    width=card_width,
    height=card_height,
)
pdf.save()