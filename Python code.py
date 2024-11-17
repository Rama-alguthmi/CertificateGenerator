# Import libraries 
import arabic_reshaper                        # To reshape Arabic text so letters connect properly
from bidi.algorithm import get_display        # To display reshaped Arabic text correctly in right-to-left (RTL) format
from reportlab.pdfgen import canvas           # To write text on a PDF page (used to add names)
from reportlab.lib.pagesizes import A4        # To set the page size to A4 for consistency
from reportlab.pdfbase.ttfonts import TTFont  # To register a custom Arabic font for writing names
from reportlab.pdfbase import pdfmetrics      # To manage and register fonts in the PDF
import os                                     # To manage folders and files (like creating folders and removing temporary files)
from PyPDF2 import PdfReader, PdfWriter       # To merge the text layer with the certificate template


  
def printCertificates(template_pdf_path, output_folder, font_path):
    # List of names to add to the certificates
    names = ["نورا احمد القرني", "محمد خالد تركستاني", "فهد احمد الحربي"]

    # Create the output folder if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Register the Arabic font for use in the certificates
    pdfmetrics.registerFont(TTFont("ArabicFont", font_path))

    for name in names:

        # Prepare Arabic text: reshape and adjust for RTL display
        reshaped_text = arabic_reshaper.reshape(name)  # Connect Arabic letters properly
        bidi_text = get_display(reshaped_text)  # Make text display correctly in RTL format

        # Create a temporary text layer PDF
        text_pdf_path = os.path.join(output_folder, f"{name}_text_layer.pdf")
        extra_width = 50  # Add extra padding for text placement
        c = canvas.Canvas(text_pdf_path, pagesize=(A4[0] + extra_width, A4[1]))

        # Set the font and size for writing the name
        font_size = 50
        c.setFont("ArabicFont", font_size)

        # Calculate where to place the text 
        text_width = c.stringWidth(bidi_text, "ArabicFont", font_size)
        x_position = ((A4[0] - text_width) / 2) + 90 + (extra_width / 2)  
        y_position = (A4[1] * 0.4) - 35  

        # Draw the name on the text layer
        c.setFillColorRGB(0.04, 0.83, 0.93)  # Set the text color 
        c.drawString(x_position, y_position, bidi_text)

        # Save the text layer
        c.save()

        # Merge the text layer with the template to create the final certificate
        final_pdf_path = os.path.join(output_folder, f"{name}.pdf")
        with open(template_pdf_path, "rb") as template_file, open(text_pdf_path, "rb") as text_layer_file:
            template_pdf = PdfReader(template_file)  # Load the template PDF
            text_layer_pdf = PdfReader(text_layer_file)  # Load the text layer PDF
            writer = PdfWriter()

            # Merge the text layer onto the template
            template_page = template_pdf.pages[0] # Access the first page of the certificate template
            text_layer_page = text_layer_pdf.pages[0] # Access the first page of the text layer (contains the name)
            template_page.merge_page(text_layer_page)  # Combine the two PDFs

            # Save the final certificate
            writer.add_page(template_page)
            with open(final_pdf_path, "wb") as output_pdf:
                writer.write(output_pdf)

        # Delete the temporary text layer PDF 
        os.remove(text_pdf_path)
        print(f"Certificate saved for {name}")

# Paths
template_pdf_path = ".../template.pdf"  # Path to the certificate template
output_folder = "..."  # Folder to save the certificates
font_path = ".../NotoSansArabic-Medium.ttf"  # Path to the Arabic font file

# Run the function
printCertificates(template_pdf_path, output_folder, font_path)
