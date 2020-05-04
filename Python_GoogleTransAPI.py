from googletrans import Translator
from text_unidecode import unidecode
import os,xlrd,xlwt
import requests
import xlrd 
session = requests.Session()
session.trust_env=False
inExcel= (r'''C:\Users\Diwakar\Desktop\Python\sample1.xlsx''')
outExcel= ("OSample.csv")
translator = Translator() 
if os.path.isfile(outExcel):os.remove(outExcel)#delete file if it exists
workbook = xlrd.open_workbook(inExcel)
sheetIn = workbook.sheet_by_index(0)

workbook = xlwt.Workbook()
sheetOut = workbook.add_sheet('DATA')
cols=[0]
for r in range(sheetIn.nrows):
	for c in cols:
									
                                    translated = translator.translate(sheetIn.cell_value(r, c), src='en',dest='hi')
                                    print(translated.text)
                                    #print(translated.pronunciation)
                                    #print(translated.extra_data['translation'])
                                    #print(translated.extra_data['translation'][0][1])
                                    #print(translator.detect(translation.extra_data['translation'][0][1]))
                                    
                                    sheetOut.write(r, c,translated.text)
                                                                        
                                    workbook.save(outExcel)#save the result
                                       
                                    
