

import xlwt        
from collections import Counter        
from nltk.corpus import stopwords
stop = set(stopwords.words('english'))
  
book = xlwt.Workbook() # create a new excel file
sheet_test = book.add_sheet('word_count') # add a new sheet
i = 0
sheet_test.write(i,0,'word') # write the header of the first column
sheet_test.write(i,1,'count') # write the header of the second column
sheet_test.write(i,2,'ratio') # write the header of the third column
    
with open('/Users/UserName/Desktop/Linkedin.rtf','r',encoding='utf-8', errors = 'ignore') as linkedindata:
     
    # convert all the word into lower cases
    # filter out stop words
    word_list = [i for i in linkedindata.read().lower().split() if i not in stop]
    word_total = word_list.__len__()
     
    count_result =  Counter(word_list)
    for result in count_result.most_common(100):
        i = i+1 
        sheet_test.write(i,0,result[0])
        sheet_test.write(i,1,result[1])
        sheet_test.write(i,2,(result[1]/word_total))
    
book.save('/Users/UserName/Desktop/WC/word_count.xls')