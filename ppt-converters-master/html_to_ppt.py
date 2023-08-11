# importing the required module
import pdfkit

# configuring pdfkit to point to our installation of wkhtmltopdf
config = pdfkit.configuration(wkhtmltopdf=r"C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe")

# converting html file to pdf file
pdfkit.from_file('text.html', 'output.pdf', configuration=config)

import aspose.slides as slides

# Create presentation
with slides.Presentation() as pres:

    # Remove default slide from presentation
    pres.slides.remove_at(0)

    # Import PDF to presentation
    pres.slides.add_from_pdf("output.pdf")

    # Save presentation
    pres.save("pdf-to-ppt.pptx", slides.export.SaveFormat.PPTX)

