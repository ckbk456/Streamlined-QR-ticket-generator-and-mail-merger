import pandas as pd
import qrcode
from pypdf import PdfWriter, PdfReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

data_sheet = "Attendance sheet.csv"
template_link = "template/ticket-template.pdf"

with open(data_sheet) as file:
    df = pd.read_csv(file, encoding="utf8")

for row in df.iterrows():
    # Column names to be adjusted
    file_name = f"0{str(row[1]['sđt'])}"
    file_content = row[1]['attendance']
    name = row[1]["họ và tên"]
    class_ = row[1]["lớp"]
    email = row[1]["email"]

    qr_image = f"temp/{file_name}.png"
    overlay_pdf_name = f"temp/{file_name}overlay.pdf"
    ticket_pdf_name = f"final/{file_name}.pdf"

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=0,
    )
    qr.add_data(file_content)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")
    print(type(img))  # qrcode.image.pil.PilImage
    img.save(qr_image)

    # Create the overlay file
    ticket = canvas.Canvas(overlay_pdf_name, pagesize=(283.465, 141.732))

    # Font for overlay file
    pdfmetrics.registerFont(TTFont('Roboto Slab', 'RobotoSlab.ttf'))

    # Draw text on overlay => manual trial and error
    text1 = canvas.PDFTextObject(ticket, x=50, y=52.5)
    text1.setFont("Roboto Slab", 6)
    text1.textLine(text=name)
    ticket.drawText(text1)

    text2 = canvas.PDFTextObject(ticket, x=50, y=40)
    text2.setFont("Roboto Slab", 6)
    text2.textLine(text=class_)
    ticket.drawText(text2)

    text3 = canvas.PDFTextObject(ticket, x=50, y=27.5)
    text3.setFont("Roboto Slab", 6)
    text3.textLine(text=file_name)
    ticket.drawText(text3)

    text4 = canvas.PDFTextObject(ticket, x=50, y=15)
    text4.setFont("Roboto Slab", 6)
    text4.textLine(text=email)
    ticket.drawText(text4)

    # Draw image on overlay
    ticket.drawInlineImage(image=img, x=156.714, y=141.732-127.559, height=113.385, width=113.385)

    # Saving overlay
    ticket.showPage()
    ticket.save()
    # Finish creating overlay file with reportlab, now on to merging overlay over template
    # Get the overlay file
    overlay = PdfReader(open(overlay_pdf_name, "rb"))

    # Load template and create Export file (PdfWriter object)
    template_pdf = PdfReader(open(template_link, "rb"))
    export = PdfWriter()

    # Putting page 0 of Overlay file on top of page 0 of Template file
    template_page = template_pdf.get_page(0)
    template_page.merge_page(overlay.get_page(0))
    # Add merged page onto Export object
    export.add_page(template_page)

    # Write Export object to disk
    with open(ticket_pdf_name, "wb") as outputStream:
        export.write(outputStream)