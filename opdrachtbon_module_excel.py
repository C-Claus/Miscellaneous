import xlsxwriter 
import random 


def create_workbook(unique_number):
    workbook = xlsxwriter.Workbook('excel_files/opdrachtbon_claus_' + unique_number + '.xlsx')   
    worksheet = workbook.add_worksheet('opdrachtbon_blad')
    bold = workbook.add_format({'bold': True})
    format = workbook.add_format({'bold': True, 'font_color': 'red'})
    
    cell_ccl = workbook.add_format()
    #cell_ccl.set_pattern(1)
    #cell_ccl.set_bg_color(chb_red)
    cell_ccl.set_font_size(9)
    cell_ccl.set_align('right')
    #cell_ccl.set_font_color('white')
    
    cell_header = workbook.add_format()
    cell_header.set_font_size(9)
    
    cell_top = workbook.add_format()
    cell_top.set_align('top')
    cell_top.set_font_size(9)
    
    #create_contents_column_A():
    worksheet.write('A1', 'CLAUS BV', bold) 
    worksheet.write('A2', 'In samenwerking met', cell_header)
    worksheet.write('A3',  'H.L. van den Boom B.V', cell_header)
    worksheet.insert_image('A4', 'image_files/claus_logo_klein.png')
    worksheet.insert_image('A9', 'image_files/hb_logo_klein.png')
    
    worksheet.write('A18', 'Opdrachtomschrijving', cell_top)
    worksheet.write('A19', 'Voorwaarden', cell_top)
    
    
    worksheet.write('A29', 'Didamseweg 6', cell_top)
    worksheet.write('A30', '6983 BB Doesburg', cell_top)
    
    worksheet.write('A32', '0313-472 475', cell_top)
    worksheet.write('A33', 'info@clausdoesburg.nl', cell_top)
    worksheet.write('A36', 'KvK 09096007', cell_top)
    
    worksheet.set_column('A:A', 20)
    
    #create_contents_column_B():
    worksheet.write('B1', 'Opdrachtnummer vermelden op afleverbon en factuur >', bold)
    worksheet.write('B4', 'Datum:', cell_ccl)
    
    worksheet.write('B6', 'Werkomschrijving:', cell_ccl)
    worksheet.write('B7', 'Werknummer:', cell_ccl)
    
    worksheet.write('B9', 'Afleveradres:',cell_ccl)
    worksheet.write('B10', 'Plaats:',cell_ccl)
    
    worksheet.write('B12', 'Leverancier:', cell_ccl)
    worksheet.write('B13', 'Contactpersoon:', cell_ccl)
    
    worksheet.write('B15', 'Leveringsdatum:', cell_ccl)
    worksheet.write('B16', 'Behandeld door:', cell_ccl)
    
    worksheet.set_row(17, 100) 
    
    
    worksheet.merge_range('B18:C18','', cell_top)
    worksheet.merge_range('B19:D19', 'De genoemde prijzen zijn exclusief btw.', cell_top)
    worksheet.merge_range('B21:D21', '*uitsluitend voor zover niet in afwijking van bovenstaande gegevens en', cell_top)
    worksheet.merge_range('B22:D22', 'of de in of via deze opdracht vermelde omschrijving en toepasselijk', cell_top)
    
    worksheet.merge_range('B23:D23', 'verklaarde voorwaarden', cell_top)
    
    worksheet.merge_range('B25:D25', 'Uitsluitend volgens de algemene voorwaarden van onderaanneming ', cell_top)
    worksheet.merge_range('B26:D26', 'AVBB / FAANB 1993 dragen wij u de werkzaamheden omschreven  ', cell_top)
    worksheet.merge_range('B27:D27', 'bij opdrachtomschrijving op.  ', cell_top)
    
    worksheet.merge_range('B30:D30', 'Betaling zal plaatsvinden 30+5 dagen na levering', cell_top)
    worksheet.merge_range('B31:D31',  'gereed werkzaamheden en ontvangst van uw factuur', cell_top)
    
    worksheet.merge_range('B33:D33', 'Uw factuur met vermelding van opdrachtnummer en werknummer te verzenden naar', cell_top)
    worksheet.merge_range('B34:D34', 'facturen@clausdoesburg.nl', cell_top)
    
    worksheet.merge_range('B36:D36', 'Statiegeld pallets dienen retour genomen te worden', cell_top)
    
    worksheet.set_column('B:B',20)
        
        
    
    
    #create_contents_column_C(): 
    worksheet.set_column('C:C', 30)
        
    #create_contents_column_D(): 
    worksheet.write('D1', str(unique_number), bold)
    
    worksheet.set_column('D:D',10)
        
        
    #create_contents_column_A()
    #create_contents_column_B()
    #create_contents_column_C()
    #create_contents_column_D()
    
    
    workbook.close()
   
for i in range(0, 100): 
    create_workbook(unique_number=str(i))
    
    