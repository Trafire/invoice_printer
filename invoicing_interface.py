import print_report, find_invoices, signatures
import os
from datetime import date, datetime
from tkinter import Tk, Label, Button, Entry, IntVar, END, W, E
from tkinter import *

class InvoicePrinter:

    def __init__(self, master):
        self.master = master
        master.title("Invoice Printer")

        self.label_invoice_num = Label(master, text="invoice:")
        self.label_date = Label(master, text="Date:")
        self.label_supplier = Label(master, text="supplier:")
        self.label_filename = Label(master, text="filename:")
        
        self.invoice_num = Entry(master)
        v = StringVar(root, value=datetime.strftime(date.today(), "%d/%m/%y"))
        self.date = Entry(master, textvariable=v)
        self.supplier = Entry(master)
        self.filename = Entry(master)

        self.run_report = Button(master, text="Run Report", command=self.run_report)
        self.print_report = Button(master, text="Print Report", command=self.print_reports)
        self.view_invoice = Button(master, text="View Invoice", command=self.view_invoice)
        self.update_invoice_list = Button(master, text="Update Invoice", command=self.update_invoice)
        self.view_report = Button(master, text="View Report", command=self.view_report)

        self.date.grid(row=1, column=1, columnspan=10, sticky=W+E)
        self.supplier.grid(row=2, column=1, columnspan=10, sticky=W+E)
        self.invoice_num.grid(row=3, column=1, columnspan=10, sticky=W+E)
        self.filename.grid(row=4, column=1, columnspan=10, sticky=W+E)

        self.label_invoice_num.grid(row=3, column=0, sticky=W)
        self.label_date.grid(row=1, column=0, sticky=W)
        self.label_supplier.grid(row=2, column=0, sticky=W)
        self.label_filename.grid(row=4, column=0, sticky=W)

        self.run_report.grid(row=5, column=0, sticky=W+E)
        self.print_report.grid(row=5, column=1, sticky=W+E)
        self.view_invoice.grid(row=5, column=2, sticky=W+E)
        self.update_invoice_list.grid(row=6, column=0, sticky=W+E)
        self.view_report.grid(row=6, column=1, sticky=W+E)

    def get_values(self):
        print_report.update_files()
        self.val_date = self.date.get().strip()
        self.val_invoice_num = self.invoice_num.get().strip()
        self.val_supplier = self.supplier.get().strip()
        self.val_filename = self.filename.get().strip().replace("\\", "\\\\")
        
    def run_report(self):
        self.get_values()
        print_report.get_report(self.val_date, self.val_supplier, self.val_filename, self.val_invoice_num)
    
    def print_reports(self):
        self.get_values()
        print_report.print_total(self.val_date, self.val_supplier, self.val_filename)

    def view_invoice(self):
        self.get_values()
        os.startfile(self.val_filename)
    def update_invoice(self):
        print_report.update_files()
        find_invoices.update(500)
    def view_report(self):
        self.get_values()
        print_report.run_distribution_report_screen(self.val_date, self.val_supplier)


root = Tk()
my_gui = InvoicePrinter(root)
root.mainloop()
