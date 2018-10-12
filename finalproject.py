################################################################
#INSTALLING AND IMPORTING THE PYTHON MODULES USED FOR THE PROJECT
###############################################################
from urllib.request import urlopen #request for url open module
from bs4 import BeautifulSoup #request for script readable with this module
import pyexcel_xls #request to export and edit the data with excel document
import re #reques the use of regex module installation
##########################################
#USAGE OF BEAUTIFUL SOUP
#########################################
def checkingtheurl():
    try:
        print("Enter a website you want to analyze with url:\n")
        url=input()
        from urllib.request import urlopen
        fetchingurl=urlopen(url)
        Soup=BeautifulSoup(fetchingurl,"html.parser")
        for script in Soup(["script","style"]):
            script.extract()
            text=Soup.get_text()
    except ValueError:
        print("The url entered is invalid !!\n")
        checkingtheurl()
    else:
        return(text)

############################################
#GETTING KEYWORDS AND FINDING DENSITY VALUES
############################################

text=checkingtheurl()#the text of all the in the URL
alllist=text.split()#all the main text has been splited with the codes
c=0#the total number of words
for x in alllist:
    c+=1
density={alllist[0]:1}# the density of each word
for x in alllist:
    cx=0
    for b in alllist:#if the words match its gets counted and upadate the density
        if(b==x):
            cx+=1
    density.update({x:cx/c})#DENSITY EQUATION
    
#########################################################
#LISTING AND SORTING THE KEY AND DENSITY VALUES IN ORDER   
##########################################################
li=list(density.values())#list of density values from top to bottom (high to low)
li.sort()
li.reverse()

#################################################################
#PUTTING INTO DICTIONARY AND ZIPPING THE VALUES WITH THE KEYWORDS
#################################################################
D={}
li2=sorted(density, key=density.__getitem__)#list of all the density values with spaces
li2.reverse()
for key,value in zip(li2,li): #ziping the both key words and the equalent density values
    D.update({key:value})
li=list(D.keys())
li2=list(D.values())

#######################################################################
#LISTING TOP FIFTEEN WORDS AND ZIP THEM TOGETHER WITH THE DENSITY VALUES
#######################################################################
i=1
print("the top fifteen words in the given url are:\n")
for x,y in zip(li,li2):
    print(x,":",y)
    i+=1
    if(i==15):
        break
word=input("enter the word for which you want to know the density of:\n")

##################################
#CREATING WORKBOOK AND WORKSHEET
##################################
worksheet=""
def writeExcelOutput(cursor):
    import xlsxwriter
    workbook=xlsxwriter.Workbook('outputexcel.xlsx')
    global worksheet
    worksheet=workbook.add_worksheet()
    for data1,data2 in zip(cursor.keys(),cursor.values()):
        j=1        
        x='B'+str(j)
        y='C'+str(j)
        z='D'+str(j)
        worksheet.write(x,data1)
        worksheet.write(y,data2*c)
        worksheet.write(z,data2)
        # number of times iteration
        if(j==16):
            j=j+1 
            break
    f=0
    
    for x in D.keys():
        if(x==word):
            f=1
    if(f==1):
        worksheet.write('B13','The details of the word entered are:') 
        worksheet.write('B14',word)
        worksheet.write('C14',D[word]*c)
        worksheet.write('D14',D[word])
    else:
        worksheet.write('B13','the entered word does not exist')
    prepareChart(workbook);
    workbook.close()
    print('Successfully wrote')
##########################
#PREPARING CHART IN EXCEL
##########################
    
#LINE CHART
def prepareChart(workbook):
    global worksheet
    chart1=workbook.add_chart({'type':'line'})
    chart1.add_series({
      
        'categories':'=Sheet1!$B$1:$B$16',
        'values':'=Sheet1!$C$1:$C$16',
       
        })
    chart1.set_title({'name':'Name vs Count'})
    chart1.set_x_axis({'name':'top 15 words'})
    chart1.set_y_axis({'name':'count'})
    
    chart1.set_style(10)#SYTLE 10 IS COLORFULL
    worksheet.insert_chart('N10', chart1)#will create chart in that area
    workbook.close()

writeExcelOutput(D)


