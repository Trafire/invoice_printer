import pickle
import win32com.client

from autof2.dailytasks import purchase_distribution
import time
import find_invoices, signatures


def get_data():
    with open("data.pic", 'rb') as data_file:
        return pickle.load(data_file)

def save_data(data):
    with open("data.pic", 'wb') as data_file:
        pickle.dump(data, data_file)
    

def get_invoice_nums(date, supplier):
    data = get_data()
    invoices = set()
    for line in data[date][supplier]:
        invoices.add(line["invoice_num"])
    text = ''
    for i in invoices:
        text += i + ", "
    return text.strip().strip(',')

def update_data(date, supplier, invoice_num, filename, distribution_report=None):
    data = get_data()
    if date not in data:
        data[date] = {}
    if supplier not in data[date]:
        data[date][supplier] = []
    record = {"invoice_num": invoice_num, "filename": filename}
    for r in data[date][supplier]:
        if r["invoice_num"] == invoice_num:
            print("here")
            if distribution_report:
                r["distribution_report"] = distribution_report
                save_data(data)
            return None
        
    if distribution_report:
        record["distribution_report"] = distribution_report
    data[date][supplier].append(record)

    save_data(data)
    

def get_report(date, supplier, filename, invoice_num):
    data = get_data()
    title = supplier + " distribution report #" + str(time.time())
    purchase_distribution.pdf_email_distribution_report(date,supplier,title = title)
    update_data(date, supplier, invoice_num, filename,title)


def update_files():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
                                    
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    data = get_data()
    for m in  find_invoices.get_distribution(messages, 25):
        for date in data:
            supplier = m[3][:6].strip()

            if supplier in data[date]:
                

                for r in data[date][supplier]:
                    if 'distribution_report' in r and r['distribution_report'] == m[3]:
                        r['distribution_filename'] = m[4]
                        save_data(data)
                        break

def print_total(date, supplier, filename):
    update_files()
    data = get_data()
    report = None
    for i in data[date][supplier]:
        if i['filename'] == filename and 'distribution_filename' in i:
            report = i['distribution_filename']
            break

    if report:
        invoices = str(get_invoice_nums(date, supplier))
        signatures.print_invoice(date, invoices, filename)
        time.sleep(.1)
        signatures.print_invoice(date, invoices, report)

def run_distribution_report_screen(date, supplier):
    purchase_distribution.run_distribution_report(date,supplier)
    
        
    
#date =  '04/04/18'
#supplier = 'CAROPR'
#filename = r"\invoices\Rosaprima International, LLC\31-03-18-Invoice #266132.pdf"
    
    
#print_total(date, supplier, filename)    
                
            
        
        
        
    
    
    
    
    
    
##    for i in data[date][supplier]:
##        if i["invoice_num"]  == invoice_num:
##            i["distribution_title"] == title
##            break
    


        
    
    
    

    

    
              


    
    
    
    



##

##data = {'04/04/18':
##        {}}
##save_data(data)





##if __name__ == "__main__":
##    invoice_num = '266282'
##    supplier = 'CAROPR'
##    date = '04/04/18'
##    filename = r"\invoices\Rosaprima International, LLC\31-03-18-Invoice #266282.pdf"
##
##
##
##    get_report(date, supplier, filename, invoice_num)
##    ##get_report(date, supplier, filename, invoice_num)
##    print_total(date, supplier, filename)

