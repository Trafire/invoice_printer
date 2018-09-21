from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
#import PIL.Image
import os
#import time

def make_print_page(dated, invoices, date):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont('Helvetica-Bold', 15)
    can.drawString(450, 100, "Recieved")
    can.drawString(450, 80, date)
    sig = ImageReader('signature.gif')
    can.drawImage(sig, 200, 60, mask='auto')
    

    if invoices:
        can.drawString(100, 100, "invoice #")
        can.setFont('Helvetica-Bold', 10)
        can.drawString(100, 80, invoices)
    can.showPage()
    can.save()

    new_pdf = PdfFileReader(packet)
    output = PdfFileWriter()
    page = new_pdf.getPage(0)
    output.addPage(page)
    
    outputStream = open("destination.pdf", "wb")
    output.write(outputStream)
    outputStream.close()

def print_invoice(date, invoices, target ):
    packet = io.BytesIO()
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont('Helvetica-Bold', 15)
    can.drawString(450, 100, "Recieved")
    can.drawString(450, 80, date)
    sig = ImageReader('signature.gif')
    can.drawImage(sig, 200, 60, mask='auto')

    if invoices:
        can.drawString(100, 100, "invoice #")
        can.setFont('Helvetica-Bold', 10)
        can.drawString(100, 80, invoices)
    can.showPage()
    can.save()

    #move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    # read your existing PDF
    existing_pdf = PdfFileReader(open(target, "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))

    output.addPage(page)

    num_pages = existing_pdf.getNumPages()
    for i in range(1,num_pages):
        output.addPage(existing_pdf.getPage(i))

    # finally, write "output" to a real file
    count = 1
    destination = r"invoices\tmp\destination%s.pdf" % count
    while os.path.isfile(destination):
        count += 1
        destination = r"invoices\tmp\destination%s.pdf" % count
        
    
    outputStream = open(destination, "wb")
    output.write(outputStream)
    outputStream.close()
    os.startfile(destination, "print")


def save_invoice(date, invoices, target):
    packet = io.BytesIO()
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFont('Helvetica-Bold', 15)
    can.drawString(450, 100, "Recieved")
    can.drawString(450, 80, date)
    sig = ImageReader('signature.gif')
    can.drawImage(sig, 200, 60, mask='auto')

    if invoices:
        can.drawString(100, 100, "invoice #")
        can.setFont('Helvetica-Bold', 10)
        can.drawString(100, 80, invoices)
    can.showPage()
    can.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    # read your existing PDF
    existing_pdf = PdfFileReader(open(target, "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))

    output.addPage(page)

    num_pages = existing_pdf.getNumPages()
    for i in range(1, num_pages):
        output.addPage(existing_pdf.getPage(i))

    # finally, write "output" to a real file
    count = 1
    destination = r"invoices\tmp\destination%s.pdf" % count
    while os.path.isfile(destination):
        count += 1
        destination = r"invoices\tmp\destination%s.pdf" % count

    outputStream = open(destination, "wb")
    output.write(outputStream)
    outputStream.close()
    return destination

'''
def invoice_sign(date, invoices_nums, target):
    make_print_page(date, invoices_nums)



    #move to the beginning of the StringIO buffer
    
    new_pdf = PdfFileReader(open(target, "rb"))
    
    # read your existing PDF
    existing_pdf = PdfFileReader(open("destination.pdf", "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)
    # finally, write "output" to a real file
    outputStream = open("destination.pdf", "wb")
    output.write(outputStream)
    outputStream.close()
'''

def make_copies(date, invoices, copies):
    make_print_page(date, invoices, date)
    for i in range(copies):
        os.startfile("destination.pdf", "print")

def invoice_printing_mode():
    target = 'true'
    date = "31/03/18"
    invoices = ''
    
    while target:
        print(date)
        print(invoices)
        target = input("target: ").strip()
        if target == "invoices":
            invoices = input("invoices").strip()
            
        elif target == "date":
            date = input("date: ")
        elif target == "blank":
            try:
                make_copies(date, invoices, 1)
            except:
                    print("failed")
        elif target == "open":
            target = input("open: ").strip()
            os.startfile(target)

        else:
            print_invoice(date, invoices, target)
            if 'y' in input("print blank: ").lower().strip():
                try:
                    make_copies(date, invoices, 1)
                except:
                    print("failed")
                    
            
        
        
                       

if __name__ == "__main__":
    invoice_printing_mode()
#target = r"\invoices\orchardcreek@cogeco.ca\30-03-18-invoice_5814.pdf"
#print_invoice(date, invoices, target)



    
