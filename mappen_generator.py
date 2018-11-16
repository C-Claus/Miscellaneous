import os
import csv
import re 

regex = re.compile('[^a-zA-Z]')

newpath = 'bibliotheek'

if not os.path.exists(newpath):
    os.makedirs(newpath)
    

with open('xml_bestanden/classificatie.csv') as csvfile:
    reader = csv.reader(csvfile, delimiter = ' ', quotechar='|')
    
    #creeert hoofdmappen
    for row in reader:
        if len(row) > 3:
            

            if row[1].isupper() == True:
                x =  str(row[1].split(';')[2] + row[1].split(';')[3]).replace('-','_')
                

            if row[1].islower() == True:
                 
     
                        
                y = row[1].split(';')[2] + '_' \
                + regex.sub('', str(row[3])) \
                + '_' + regex.sub('', str(row[4]))\
                + '_' + regex.sub('', str(row[5]))\
                + '_' + regex.sub('', str(row[6])\
                #+ '_' + regex.sub('', str(row[7]) \
                
                ) 
            
                
                if 'Tabel' not in y:
                    print y
                    
                    z = y.replace('NLSfB', '')
                    
                    print z 
                    
                    
                    os.makedirs(newpath + '/' + x + '/' + z)
                
                
    