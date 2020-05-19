# -*- coding: utf-8 -*-
"""
Created on Fri Oct 18 18:32:16 2019

@author: Edward
"""

#from PIL import Image 
#import PyPDF2
import os 
import re
#from glob import glob
#import csv
import difflib
import pytesseract
import pandas as pd
#pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Edward\AppData\Local\Programs\Python\Python37\Lib\site-packages\tesseract.exe'

working_dir = os.getcwd()

pytesseract.pytesseract.tesseract_cmd = working_dir + '\\tesseract_x64-windows-static\\tools\\tesseract\\tesseract.exe'

dir1 =  working_dir + '\\poppler-0.68.0\\bin'

import pdf2image as p2
import shutil


import numpy.random.common
import numpy.random.bounded_integers
import numpy.random.entropy

#working_dir = 'D:\\Employment\\SINGER AC\\Scanned'

os.chdir(working_dir)
os.getcwd()

customer_list = ['A.C SCHULTESNC',
 'Acuatubos',
 'Advance Engineering',
 'AELENEVA MEXICO',
 'Alianca S.A. Ind Naval',
 'AMERICAN WATER WORKS',
 'Aqua Works',
 'Artic Control',
 'Assistech',
 'Astro Maquinaria Corp.',
 'AVK Valvulas',
 'B & C Exports',
 'Brigitte Campbell',
 'Capitol Pump Resources',
 'Caspin Water',
 'Central water & sewerage authority',
 'Chamco Industries Ltd.',
 'Cimco Sales and Marketing',
 'COASTAL PROCESS',
 'CORE AND MAIN LP',
 'Dao Nguyen',
 'Daybreak Technologies',
 'Delpin',
 'DELTA T EQUIP/ TOTAL PUMPS',
 'Durga',
 'Dynamic Process',
 'E.P.C. Engineering Products',
 'East Asia',
 'ELECTRIC PUMP INC',
 'Empresa Mexicana De Manufacturas SA',
 'ENGINEERED FLUIDS INC',
 'Engineered Systems',
 'Euro Oriental trading co. Ltd.',
 'Fareco',
 'FERGUSON WATERWORK',
 'FIMALI LIMITED',
 'Fire Lion global LLC',
 'FIRST SUPPLY',
 'FLOCOR INC',
 'FLORIDA VALVE',
 'FLOSOURCE',
 'FLOTECH',
 'FLUICONST',
 'FORTLILINE',
 'FRONTIER SUPPLY',
 'Garcia Llerandi',
 'Golf Pumping Services',
 'G-Tech',
 'Henry Pratt',
 'HMA Flow',
 'INCOTEK COMPANY LTD',
 'INSTRUMENT and SUPPLY WEST',
 'Isaacs & Associates, Inc',
 'IWIAPRODUCTOS C.A.',
 'James Electric motor services',
 'James, Cooke & Hobson',
 'JG ACUEDUCTOS',
 'JINGMEN PRATT VALVE',
 'KENNEDY INDUSTRIES',
 'Kiho USA',
 'KURODA NORTE',
 'M.L.K. AND ASSOCIATES',
 'Marketing M & E',
 'Metaval',
 'MGA Controls',
 'Mikala Enterprises Limited',
 'MIYA WATER PRODUCTS',
 'MONTSERRAT UTILITIES',
 'mueller Company, Decatur, IL',
 'Mueller Middle East',
 'MUNICIPAL EQUIP',
 'NATIONAL WATER & SEWERAGE',
 'Omnitech',
 'Pipestone',
 'POWERCODE ENTERPRISE',
 'Prochem',
 'Provan',
 'PT Pancatama',
 'R&D Fire Products',
 'R.W. PROPERTY OWNERS',
 'Rich Klinger',
 'Rocky Mountain',
 'Rola',
 'Ruhrpumpen',
 'Ruhrpumpen Systems',
 'SAGIT FZE',
 'Singer Valve LLC',
 'Singer Taicang',
 'Smith Tech',
 'Southwest Valve',
 'SPARTAN CONTROLS',
 'STRAEFFER PUMP',
 'Sullair',
 'Summit Valve',
 'SVM',
 'Syntec',
 'Tavira',
 'TEK IN LLAVENMAG',
 'TEXAS FLUID POWER',
 'THE BROWN COMPANY',
 'The Gellert',
 'TIGERFLOW',
 'Total Flow Control',
 'TRI-STATE IND',
 'Tydan ',
 'Valveco',
 'VESSCO INC',
 'Viet An Environment Technology',
 'VIRGIN ISLANDS',
 'Virgin Valley',
 'Water works suppliers',
 'Waterpoint']

customer_list = [x.lower() for x in customer_list]


output_dir = working_dir + '\\renamed'
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

pdf_list = []
xls_list = []
doc_list = []

for file_name in os.listdir(working_dir):
    if re.search('(.pdf)$' , file_name.lower()):
        pdf_list.append(file_name)
    if re.search('(xls)$' , file_name.lower()):
        xls_list.append(file_name)
    if re.search('(doc)$' , file_name.lower()):
        doc_list.append(file_name)
 
###PDF


print(len(pdf_list), "PDF files found in this directory")

'''
########This code for files that are already ORCed into rich text format
    pdf_file = open(file, mode="rb")
    read_pdf = PyPDF2.PdfFileReader(pdf_file)
    pageObj = read_pdf.getPage(0)
    mystring = str(pageObj.extractText()).lower()
'''


SN_finder_JDE = re.compile('([0-9]{7,7})\s*(so|xo|ms|xr|mr|sr)')
SN_finder_syspro = re.compile('([0-9]{8,8})')


def find_jde_SN(pages):
    for im in pages:    
        width, height = im.size 
        left = width*1/3
        top = height/10 - 100
        right = width*2/3 + 100
        bottom = height*2 / 10 + 100
        im1 = im.crop((left, top, right, bottom))    
        
        mystring = str(pytesseract.image_to_string((im1))).replace('\n', '').lower()
        
        if bool(SN_finder_JDE.search(mystring)) == True:
            found_group = SN_finder_JDE.findall(mystring)
            if len(found_group) ==1:
                output = found_group[0][0] +' ' + found_group[0][1].upper()
                certainty = 1
                break
            elif len(found_group) >=1:
                for found in found_group:
                    something_wrong = True
                    if (found_group.count(found) / len(found_group)) == 1:
                        output = found[0] + ' ' + found[1].upper()     
                        certainty = 1
                        something_wrong = False
                    elif (found_group.count(found) / len(found_group)) > 0.6:
                        output = 'This SN is likly: ' + found[0] + ' ' + found[1].upper() + ', but there were multiple matches'
                        certainty = 2
                        something_wrong = False
                if something_wrong == True:
                    output = 'Something is wrong, more than one matching SN'    
                    certainty = 3
                break
        else:
            output = 'Did not find a matching SN'
            certainty = 4
    return (certainty, output)



def find_syspro_SN(pages):
    for im in pages:    
        width, height = im.size 
        left = width - width /4
        top = 0
        right = width
        bottom = height / 10
        im1 = im.crop((left, top, right, bottom)) 
        
        mystring = str(pytesseract.image_to_string((im1))).replace('\n', '').lower()
        
        if bool(SN_finder_syspro.search(mystring)) == True:     
            found_group = SN_finder_syspro.findall(mystring)
            if len(found_group) ==1:
                output = found_group[0]   
                certainty = 1
                break
            elif len(found_group) >=1:
                for found in found_group:
                    something_wrong = True
                    if (found_group.count(found) / len(found_group)) == 1:
                        output = found
                        certainty = 1
                        something_wrong = False
                    elif (found_group.count(found) / len(found_group)) > 0.6:
                        output = 'This SN is likly: ' + found + ', but there were multiple matches'
                        certainty = 2
                        something_wrong = False
                if something_wrong == True:
                    output = 'Something is wrong, more than one matching SN'
                    certainty = 3
                break
        else:
            output = 'Did not find a matching SN'
            certainty = 4
    return (certainty, output)




#########################################################
'''test'''
# pages = p2.convert_from_path(pdf_list[1], 500, single_file = True, poppler_path = dir1) 
# im = pages[0]

# width, height = im.size 
# left = 0
# top = height/8
# right = width/2 + 50
# bottom = height / 4 + 100
# im1 = im.crop((left, top, right, bottom)) 
# im1    

# # from pytesseract import Output
# # import cv2
# # d = pytesseract.image_to_data(im1, output_type=Output.DICT)
# # d.keys()
# # d['left']

# # n_boxes = len(d['level'])
# # for i in range(n_boxes):
# #     (x, y, w, h) = (d['left'][i], d['top'][i], d['width'][i], d['height'][i])
# #     cv2.rectangle(im1, (x, y), (x + w, y + h), (0, 255, 0), 2)
 

# mystring = str(pytesseract.image_to_string((im1))).lower()

# splits = re.split(': |, |\*|\n',mystring)

# output = 'Did not find a matching Company'
# for word in splits:
#     best_match = difflib.get_close_matches(word, customer_list)
#     if len(best_match) >= 1:
#         output = best_match[0]
#         break

# print (output)

'''test'''
#################################################################



def find_JDE_company(pages):
    output = 'Did not find a matching Company'
    for im in pages:    
       width, height = im.size 
       left = 0
       top = height/8
       right = width/2 + 50
       bottom = height / 4 + 100
       im1 = im.crop((left, top, right, bottom)) 
       im1    
       
       mystring = str(pytesseract.image_to_string((im1))).lower()
       splits = re.split(': |, |\*|\n',mystring)
       
       for word in splits:
           best_match = difflib.get_close_matches(word, customer_list)
           if len(best_match) >= 1:
               output = best_match[0]
               break

    return output





def copy_move (file, sn, company):
    old_name = working_dir +'\\' + file
    new_name = output_dir + '\\' + company + ' - ' + sn +'.pdf'
    shutil.copy(old_name, new_name)
  

SN_dict = {}
company_dict = {}
                         
def main():

    exact_match = 0
    counter = 1
    for file in pdf_list:
        
        print('Searching PDF ', counter, ' out of ', len(pdf_list))
        
        #OCR the first page only, (single_file = True)
        pages = p2.convert_from_path(file, 500, single_file = True, poppler_path = dir1) 
        
        jde_output = find_jde_SN(pages)
        SN_found = False
       
        if jde_output[0] == 1:
            output = jde_output[1]
            exact_match += 1
            SN_found = True

        else:   
            syspro_output = find_syspro_SN(pages)
            if syspro_output[0] == 1:
                output = syspro_output[1]
                exact_match +=1
                SN_found = True
            else:
                if jde_output[0] <= syspro_output[0]:
                    output = jde_output[1]
                else:
                    output = syspro_output[1]
    
        SN_dict[file] = output
        company_dict[file] = find_JDE_company(pages).title()
        counter +=1
        
        if SN_found == True and company_dict[file] != 'Did Not Find A Matching Company':
            copy_move(file, output, company_dict[file])
        
    
    SN_df = (pd.DataFrame(data = SN_dict, index =[0]).T)
    SN_df.columns = ['SN# Result']
    
    company_df = (pd.DataFrame(data = company_dict, index =[0]).T)
    company_df.columns = ['Company Result']
    
    df = pd.concat([SN_df, company_df], axis=1, sort=False)

    writer = pd.ExcelWriter(output_dir + '\\reference_list.xlsx')
    df.to_excel(writer)
    writer.save()
    
    
    
    '''Out Put message'''
    print('Number of original files:', len(pdf_list))
    print('Number of exact SN matches:', exact_match)
    print('See \'reference_list.xlsx\' in the \'renamed\' folder for detail')
    print('FINISHED')
    

if __name__ == '__main__':
    main()

# pdf_file = open(pdf_path, mode="rb")
# read_pdf = PyPDF2.PdfFileReader(pdf_file)
# pageObj = read_pdf.getPage(0)
# mystring = str(pageObj.extractText()).lower()

        
# SN_finder_JDE = re.compile('([0-9]{7,7})\s(so|xo|ms|xr|mr|sr)')
# found_group = SN_finder_JDE.findall(mystring)
# found_group


# found_range = SN_finder_JDE.search(mystring)
# found_range


# SN_finder_syspro = re.compile('\s([0-9]{8,8})\s')
# found_group = SN_finder_syspro.findall(mystring)
# found_group

# found_range = SN_finder_syspro.search(mystring)
# found_range


# """
# for j in range(0, len(filenamelist)):
#     #j=0
#     pdf_file = open(filenamelist[j], mode="rb")
#     read_pdf = PyPDF2.PdfFileReader(pdf_file)
#     pageObj = read_pdf.getPage(0)
#     print(pageObj.extractText()) 
    
#     mystring = str(pageObj.extractText())
    
#     for i in range(0, len(mystring) - 7):
#         SNcheck = mystring[i:i+7]
#         if re.match('^[0-9]{7}',SNcheck) is not None:
#             if mystring[i+8:i+10] == 'so':
#                 print (SNcheck)
#                 found = i
#                 print(i)
                
#     pdf_file.close() 
    
#     #['A.C SCHULETS OF CAROLINA-3528071SO-OCR.pdf',
#      #'A.C SCHULTES OF CAROLINA-2519122SO-OCR.pdf']
    
#     filename = filenamelist[j]
#     SNcheck = mystring[found:found+7]
#     os.rename(filename, SNcheck + '.pdf')
#     newfilelist.append(SNcheck)
# #tab until here for loop
# """



# ###word DOC
# import textract
# #for file in doc_list:
# file = doc_list[1]

# text = textract.process('D:\\Employment\\SINGER AC\\Scanned\\' +file)


# doc_file = file
# docx_file = doc_file + 'x'

# os.system('antiword ' + doc_file + ' > ' + docx_file)
# text = open(docx_file).read()
# os.remove(docx_file)




# doc_file = open(file, mode="rb")
# read_pdf = PyPDF2.PdfFileReader(pdf_file)
# pageObj = read_pdf.getPage(0)
# mystring = str(pageObj.extractText())
    
# for i in range(10, len(mystring) - 7):
#     SNcheck = mystring[i:i+8]
#     if re.match('^[0-9]{8}',SNcheck) is not None:
#          if mystring[i-10:i-1] == 'SHOP COPY':
#             print (SNcheck)
#                 #found = i
#                 #print(i)
                
# pdf_file.close() 


# #add sold to name, prepare for different format of Sales order.
# #How many formats are there??