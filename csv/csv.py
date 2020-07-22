import xlrd
import xlwt
from collections import Counter 
book=xlrd.open_workbook(r"C:\Users\technogeek\Desktop\csv\Book2.xlsx")
print(book.nsheets)
print(book.sheet_names())
arshad=book.sheet_by_index(0)

print(arshad.nrows)
print(arshad.ncols)
for i in range(0,arshad.nrows):
    for j in range(0,arshad.ncols):
        
       print(arshad.cell(i,j))


#opening secondfile 
book2=xlrd.open_workbook(r"C:\Users\technogeek\Desktop\csv\Book1.xlsx")
print(book2.nsheets)
print(book2.sheet_names())
arshad2=book2.sheet_by_index(0)
print(arshad2.nrows)
print(arshad2.ncols)

for i in range(0,arshad2.nrows):
    for j in range(0,arshad2.ncols):
        
       print(arshad2.cell(i,j))


#000000xgfssj 

if(arshad.nrows<arshad2.nrows):
    big_rows=arshad2.nrows
elif(arshad2.nrows<arshad.nrows):
    big_rows=arshad.nrows
else:
    big_rows=arshad.nrows


print(big_rows)

if(arshad.ncols<arshad2.ncols):
    big_col=arshad2.ncols
elif(arshad2.ncols<arshad.ncols):
    big_col=arshad1.ncols
else:
    big_col=arshad.ncols

print(big_col)

#-----------------------------------email is identifier here
email_set_data1=[]
email_set_data2=[]
for i in range(1,big_rows):
    email1=arshad.cell(i,1)
    email2=arshad2.cell(i,1)
    email_set_data1.append(email1)
    email_set_data2.append(email2)
    
print(email_set_data1)
print(email_set_data2)
res = [ ele for ele in email_set_data1] 
for i in email_set_data2: 
  if i in email_set_data1: 
    res.remove(i) 
  
# printing result 
print("The Subtracted list is : " + str(res)) 

#not it will check data i mean compare from the second sheet if the item exist in second sheet it will not add there


rtcsv=xlwt.Workbook(encoding="utf-8")
csvdata=rtcsv.add_sheet("arsh@lib")
for i in range(0,len(res)):
    csvdata.write(i,0,str(res[i]))

rtcsv.save("result.xls")

#------------------copyright information
"""Copyright <2020> <MOHAMMAD ARSHAD>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

End license text."""