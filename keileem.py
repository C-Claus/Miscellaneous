
#https://stackoverflow.com/questions/34837707/how-to-extract-text-from-a-pdf-file
#https://pymupdf.readthedocs.io/en/1.19.3/tutorial.html#opening-a-document


import os
import fitz 
import xlsxwriter



def get_text_from_pdf(pdf_file):

    workbook = xlsxwriter.Workbook(str(pdf_file.replace('.pdf','.xlsx')))
    worksheet = workbook.add_worksheet()

    worksheet.write(0,0, 'NR')
    worksheet.write(0,1, 'Losdatum')
    worksheet.write(0,2, 'Laaddatum')
    worksheet.write(0,3, 'KG')
    

    doc = fitz.open(pdf_file) 
    
    list_nummer = []
    list_laaddatum = []
    list_losdatum = []
    list_kg = []
    total_list = []

    
    for page_number in range(0, int(doc.page_count)):

        page = doc.load_page(page_number)
        text = page.get_text("words")

        for index, words in enumerate(text):
            
            if 'NR' in words:
                index_nr = index

            if 'Laaddatum' in words:
                index_laaddatum = index

            if 'Losdatum' in words:
                index_losdatum = index 

            if 'Netto' in words:
                index_netto = index


        for i, words in enumerate(text):
            if i in range(index_nr+2, index_nr+3, 1):
                list_nummer.append(words[4])

            if i in range(index_laaddatum+2, index_laaddatum+3,1):
                list_laaddatum.append(words[4])

            if i in range(index_losdatum+2, index_losdatum+3,1):
                list_losdatum.append(words[4])

            if i in range(index_netto+2, index_netto+3,1):
                list_kg.append(words[4])

 
    total_list.append(list_nummer)
    total_list.append(list_laaddatum)
    total_list.append(list_losdatum)
    total_list.append(list_kg)

    for row, item_nr in enumerate(list_nummer):
        worksheet.write(row+1, 0, item_nr)

    for row, laaddatum in enumerate(list_laaddatum):
        worksheet.write(row+1, 1, laaddatum)

    for row, losdatum in enumerate(list_losdatum):
        worksheet.write(row+1, 2, losdatum )

    for row, kg in enumerate(list_kg):
        worksheet.write(row+1, 3, kg)

    workbook.close()


pdf_formulieren = os.listdir("C:\\Algemeen\\07_prive\\keileem\\PDF\\formulieren")

for pdf_file in pdf_formulieren:

    path_folder = "C:\\Algemeen\\07_prive\\keileem\\PDF\\formulieren\\"
    path_file = path_folder + pdf_file

    print (path_file)

    get_text_from_pdf(pdf_file=str(path_file))
