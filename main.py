import fpdf
import openpyxl
import PyPDF2
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
from tkinter import *
from tkinter import filedialog

def load_excel():
    global excelpath
    excelpath = filedialog.askopenfile(filetypes = [("Excel files", "*.xlsx"), ("Excel files", "*.xlsm")])
    ex_info.set(excelpath.name)
    label_message.set("")

def load_pdf():
    global pdfpath
    pdfpath = filedialog.askopenfile(filetypes = [("PDF files", "*.pdf")])
    pdf_info.set(pdfpath.name)
    label_message.set("")

def write_data_to_pdf():
    label_message.set("")
    try:
        # loading excel file
        wb = openpyxl.load_workbook(excelpath.name)
        sheet = wb[wb.sheetnames[0]]
        num_g = sheet.max_row

        # creating dictionary with data from excel, deleting useless text
        pairs = {}
        project_num = project_entry.get()
        
        for i in range(1,num_g+1):
            if sheet[f"A{i}"].value.replace(str(project_num), "").replace("/00", "").replace("/0", "").replace("/", "").replace(" ", "") not in pairs:
                pairs[sheet[f"A{i}"].value.replace(str(project_num), "").replace("/00", "").replace("/0", "").replace("/", "").replace(" ", "")] = (sheet[f"B{i}"].value)
            elif isinstance(pairs[sheet[f"A{i}"].value.replace(str(project_num), "").replace("/00", "").replace("/0", "").replace("/", "").replace(" ", "")], list):
                pairs[sheet[f"A{i}"].value.replace(str(project_num), "").replace("/00", "").replace("/0", "").replace("/", "").replace(" ", "")].append(sheet[f"B{i}"].value)
                # deleting duplicates in each list
                pairs[sheet[f"A{i}"].value.replace(str(project_num), "").replace("/00", "").replace("/0", "").replace("/", "").replace(" ", "")] = list(dict.fromkeys(pairs[sheet[f"A{i}"].value.replace(str(project_num), "").replace("/00", "").replace("/0", "").replace("/", "").replace(" ", "")]))
                # sorting each list
                pairs[sheet[f"A{i}"].value.replace(str(project_num), "").replace("/00", "").replace("/0", "").replace("/", "").replace(" ", "")].sort()
            else:
                pairs[sheet[f"A{i}"].value.replace(str(project_num), "").replace("/00", "").replace("/0", "").replace("/", "").replace(" ", "")] = [pairs[sheet[f"A{i}"].value.replace(str(project_num), "").replace("/00", "").replace("/0", "").replace("/", "").replace(" ", "")], (sheet[f"B{i}"].value)]



        # creating list of page numbers (keys) to determine max page number
        keys = []
        for i in pairs.keys():
            keys.append(int(i))
        max_page = max(keys)

        # creating dictionary with page sizes
        existing_pdf = PdfFileReader(open(pdfpath.name, "rb"))
        count = existing_pdf.numPages
        sizes = {}
        for i in range(count):
            x = existing_pdf.getPage(i)
            sizes[i+1] = [int(x["/MediaBox"][2]), int(x["/MediaBox"][3])]

        # setting global properties of temporary PDF file and font
        temp_pdf = fpdf.FPDF(orientation = "P", unit = "pt")
        temp_pdf.set_font("Courier")
        temp_pdf.set_margins(0, 0, 0)


        doc_num = doc_entry.get()

        # iterating through dictionary and adding blank or filled pages to temporary PDF file
        for i in range(0,max_page):
            page_w = sizes[i+1][0]
            page_h = sizes[i+1][1]
            cell_w = 0.18
            space_left = 0.08
            if pairs.get(f"{i+1}") == None:
                temp_pdf.add_page(format = (page_w, page_h))
            elif type(pairs[f"{i+1}"]) is str:
                temp_pdf.set_font("", "B", size=int(page_h/33.1))
                temp_pdf.set_text_color(255, 0, 0)
                temp_pdf.set_draw_color(255, 0, 0)
                temp_pdf.set_fill_color(255, 255, 255)
                temp_pdf.set_line_width(3)
                temp_pdf.add_page(format = (page_w, page_h))
                temp_pdf.cell(int((1-space_left-cell_w)*page_w), int((page_h)/36), ln=0)
                temp_pdf.cell(int(cell_w*page_w), int((page_h)/36), txt=str(pairs[f"{i+1}"]), ln=0, border=1, fill=1, align="C")
                if doc_num == "":
                    pass
                else:
                    temp_pdf.cell(int(0.2*space_left*page_w), int((page_h)/36), ln=0)
                    temp_pdf.set_font("", "B", size=int(page_h/50))
                    temp_pdf.set_text_color(26, 26, 255)
                    temp_pdf.set_draw_color(26, 26, 255)
                    temp_pdf.set_line_width(2)
                    temp_pdf.cell(int(0.8*space_left*page_w), int((page_h)/45), txt=f"DOK-{doc_num}", ln=0, border=1, fill=1, align="C")
            else:
                temp_pdf.set_font("", "B", size=int(page_h/33.1))
                temp_pdf.set_text_color(255, 0, 0)
                temp_pdf.set_draw_color(255, 0, 0)
                temp_pdf.set_fill_color(255, 255, 255)
                temp_pdf.set_line_width(3)
                temp_pdf.add_page(format = (page_w, page_h))
                if len(pairs[f"{i+1}"]) == 2:
                    temp_pdf.cell(int((1-space_left-(2*cell_w))*page_w), int((page_h)/36), ln=0)
                    temp_pdf.cell(int(cell_w*page_w), int((page_h)/36), txt=str((pairs[f"{i+1}"])[0]), ln=0, border=1, fill=1, align="C")
                    temp_pdf.cell(int(cell_w*page_w), int((page_h)/36), txt=str((pairs[f"{i+1}"])[1]), ln=0, border=1, fill=1, align="C")
                elif len(pairs[f"{i+1}"]) == 3:
                    temp_pdf.cell(int((1-space_left-(3*cell_w))*(page_w)), int((page_h)/36), ln=0)
                    temp_pdf.cell(int(cell_w*page_w), int((page_h)/36), txt=str((pairs[f"{i+1}"])[0]), ln=0, border=1, fill=1, align="C")
                    temp_pdf.cell(int(cell_w*page_w), int((page_h)/36), txt=str((pairs[f"{i+1}"])[1]), ln=0, border=1, fill=1, align="C")
                    temp_pdf.cell(int(cell_w*page_w), int((page_h)/36), txt=str((pairs[f"{i+1}"])[2]), ln=0, border=1, fill=1, align="C")
                if doc_num == "":
                    pass
                else:
                    temp_pdf.cell(int(0.2*space_left*page_w), int((page_h)/36), ln=0)
                    temp_pdf.set_font("", "B", size=int(page_h/50))
                    temp_pdf.set_text_color(26, 26, 255)
                    temp_pdf.set_draw_color(26, 26, 255)
                    temp_pdf.set_line_width(2)
                    temp_pdf.cell(int(0.8*space_left*page_w), int((page_h)/45), txt=f"DOK-{doc_num}", ln=0, border=1, fill=1, align="C")

        # saving temporary PDF file
        temp_pdf.output("temp.pdf")

        # opening existing file, counting number of pages
        existing_pdf = PdfFileReader(open(pdfpath.name, "rb"))
        with open("temp.pdf", "rb") as f:
            temp = PdfFileReader(f)
            count = existing_pdf.numPages
            output = PdfFileWriter()

            # iteratin through pages, merging pages from two PDF files
            if count == max_page:
                for i in range(0,count):
                    page = existing_pdf.getPage(i)
                    page.mergePage(temp.getPage(i))
                    output.addPage(page)
            else:
                for i in range(0,max_page):
                    page = existing_pdf.getPage(i)
                    page.mergePage(temp.getPage(i))
                    output.addPage(page)
                for i in range(max_page, count):
                    page = existing_pdf.getPage(i)
                    output.addPage(page)

            # saving output file
            name = os.path.splitext(os.path.basename(pdfpath.name))[0]
            with open(f"{name}-desc.pdf", "wb") as outputStream:
                output.write(outputStream)

        # deleting temporary PDF file
        os.remove("temp.pdf")
        label_message.set(f"File {name}-desc.pdf was created successfully!")
    
    # handling errors
    except ValueError:
        label_message.set("Insert correct project number.")
    except NameError:
        label_message.set("At least one of the files is not loaded.")
    except openpyxl.utils.exceptions.InvalidFileException:
        label_message.set("Please load *.xlsx file.")
    except PyPDF2.utils.PdfReadError:
        temp_pdf.close()
        os.remove("temp.pdf")
        label_message.set("Please load *.pdf file.")
    except PermissionError:
        temp_pdf.close()
        os.remove("temp.pdf")
        label_message.set(f"File {name}-desc.pdf is open. Cannot overwrite.")
    except:
        label_message.set("Something went wrong :(")


root = Tk()
root.geometry("720x130")
root.title("yPDF")
root.resizable(False, False)

label_message = StringVar()
ex_info = StringVar()
pdf_info = StringVar()
font_color = StringVar()

FONT = ("Arial", 10)

frame1 = Frame(root, width=400, height=28)
frame1.pack()
frame1.pack_propagate(0)


label_num = Label(frame1, text="Project number:", font=FONT, width=12)
label_num.pack(side=LEFT)

project_entry = Entry(frame1, font=FONT, width=12)
project_entry.pack(side=LEFT)

doc_entry = Entry(frame1, font=FONT, width=5)
doc_entry.pack(side=RIGHT)

label_doc = Label(frame1, text="Doc number:", font=FONT, width=12)
label_doc.pack(side=RIGHT)


frame2 = Frame(root)
frame2.pack(pady=5)

ex_button = Button(frame2, text="Load Excel file", font=FONT, command=load_excel, width=20)
ex_button.grid(row=1, column=0, padx=5)

label_ex = Label(frame2, textvariable=ex_info, font=FONT, borderwidth=2, relief="groove", width=60, height=1)
label_ex.grid(row=1, column=1, padx=5)

pdf_button = Button(frame2, text="Load PDF file", font=FONT, command=load_pdf, width=20)
pdf_button.grid(row=2, column=0, padx=5)

label_pdf = Label(frame2, textvariable=pdf_info, font=FONT, borderwidth=2, relief="groove", width=60, height=1)
label_pdf.grid(row=2, column=1, padx=5)

data_button = Button(frame2, text="Extract data to PDF", font=FONT, command=write_data_to_pdf, width=20)
data_button.grid(row=3, column=0, padx=5)

label_info = Label(frame2, textvariable=label_message, font=("Arial", 10, "bold"), width=60, height=1)
label_info.grid(row=3, column=1, padx=5)

root.mainloop()