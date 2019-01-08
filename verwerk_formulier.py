

import xlsxwriter
import xlrd
import unicodedata
import re


excel_file_name = "U:\\75. Woonstad Rotterdam\\31-10-2018\\BIM formulier Woonstad Rotterdam(1-5).xlsx"
    
workbook = xlrd.open_workbook(excel_file_name)
worksheet = workbook.sheet_by_name("Sheet1") 

num_rows = worksheet.nrows #Number of Rows
num_cols = worksheet.ncols #Number of Columns


workbook_to_write  = xlsxwriter.Workbook("U:\\75. Woonstad Rotterdam\\31-10-2018\\formulier_verwerkt.xlsx")
worksheet_write = workbook_to_write.add_worksheet()


total_list=[]  
 
for i in range(1, num_rows, 1):
    for current_column in range(0, num_cols, 1):
        total_list.append([current_column, worksheet.cell_value(i, 4), worksheet.cell_value(i,5), worksheet.cell_value(0, current_column), worksheet.cell_value(i, current_column)])
        
        
i = 0

write_total_list=[]
pivot_table_list = []
for v in total_list:
    if isinstance(v[4], unicode):
        for tok in v[4].split(';'):
            if len(tok) > 0:
                pivot_table_list.append([i, 1, v[1], v[2],v[3], tok])
            i += 1 
            
  
header_list = ["E-mail","Rol","Vraag","Antwoord"]    

for i, v in enumerate(header_list):
    worksheet_write.write(0, i, v)      
  
width = 40     


for i in range(0, num_rows):
    worksheet_write.set_column(i,1, width)
    


#er zijn maar vier kolommen, dus hopelijk gaat dit goed

for i,v in enumerate(pivot_table_list):
    
    #E-mail
    worksheet_write.write(i+1,0, v[2])
    
    #Rol
    worksheet_write.write(i+1,1, v[3])
    
    #Vraag
    worksheet_write.write(i+1,2, v[4])
    
    #Antwoord
    worksheet_write.write(i+1,3, v[5])
    
    
           
workbook_to_write.close()             

            
            
            
        

       
