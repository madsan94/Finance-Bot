import requests
from bs4 import BeautifulSoup
import xlwt
from xlwt import Workbook
import json
import time
import sys
#-------------------------------------------Necessary Data Structures---------------------------------------------------

PROFIT_LOSS={"SALES":'',"EXPENSES":'',"OPERATING_PROFIT":'',"OPM_PER":'',"OTHER_INCOME":'',"INTEREST":'',"DEPRICIATION":'',"PBT":'',"TAX_PER":'',"NET_PROFIT":'',"EPS":'',"DIVIDEND_PAYOUT":''}
BALANCE_SHEET={"SHARE_CAPITAL":'',"RESERVES":'',"BORROWINGS":'',"OTHER_LIABILITIES":'',"TOTAL_LIABILITIES":'',"FIXED_ASSETS":'',"CWIP":'',"INVESTMENTS":'',"OTHER_ASSETS":'',"TOTAL_ASSETS":''}
CASHFLOW={"OPERATING_ACTIVITY":'',"INVESTING_ACTIVITY":'',"FINANCING_ACTIVITY":'',"NETFLOW":''}

#-------------------------------------------END of Data Structure Declaration-------------------------------------------


#-------------------------------------------Functions to format the screen and excel Output-----------------------------

#Function to Remove td tags and commas
def trim(a):
    index=a.index('/')
    final_1=a[4:index-1]
    final=final_1
    i=0
    while(i<len(final_1)):
        if(final_1[i]==','):
            final=final_1[0:i]+final_1[i+1:]
            break
        i+=1
    j=0
    while(j<len(final_1)):
        if(final_1[j]=='%'):
            final=final_1[0:j]+final_1[j+1:]
            break
        j+=1
    return(final)

#Function to return a list of data for all the years for a component of the financial statement
def Year_wise_component_breakup(C):
    list=[]
    i=1
    while(i<len(C)):
        temp=str(C[i])[0:4]
        if(temp=="<td>"):
            list.append(trim(str(C[i])))
        i+=1
    return(list)

def excel_write(A,wb,sheet,excel_name):
    k=1
    start=2007
    while(k<13):
        sheet.write(k,0,str(start))
        start+=1
        k+=1
    P_L_Keys=[]
    i=0
    j=0
    for a in list(A.keys()):
        P_L_Keys.append(a)
    while(i<len(P_L_Keys)):
        sheet.write(0,i+1,P_L_Keys[i])
        b=1
        while(b<=len(list(A[P_L_Keys[i]]))):
            sheet.write(b,i+1,(list(A[P_L_Keys[i]]))[b-1])
            b+=1
        i+=1
    excel_name="NBFC Data/"+excel_name+'.xls'
    wb.save(excel_name)

def GoogleSearch(i):
    query="https://www.googleapis.com/customsearch/v1?key=AIzaSyCNWaoOioDDJHv4fX1XUk53z_Qh1Vw9ED8&cx=000482377268081262274:6th5djadup4&q="
    query=query+i
    page=requests.get(query)
    #print(page.text)
    print(json.loads(page.text)['items'][0]['link'])
    return(json.loads(page.text)['items'][0]['link'])

def FindPage(link):
    page=requests.get(link)
    soup=BeautifulSoup(page.content,'html.parser')
    table=soup.find_all('table',class_='data-table')
    return(table)






#-----------------------------------------Writing to Excel and save--------------------------------------------------

# Creating Sheets in Excel and taking data from File
def populating_excel(x):
    file=open("NBFC Companies","r")
    #Google Search and Finding the Link
    link=GoogleSearch(x)
    table=FindPage(link)
    wb = Workbook()
    sheet1=wb.add_sheet('Profit and Loss')
    sheet2=wb.add_sheet('Balance Sheet')
    sheet3=wb.add_sheet('Cashflow')
    sheet1.write(0,0,"YEAR")
    sheet2.write(0,0,"YEAR")
    sheet3.write(0,0,"YEAR")
    #A->table[x]->
    #x=0:Peer Comparison
    #x=1:Quartely Result
    #x=2:Profit and Profit
    #x=3:Balance Sheet
    #x=4:CashFlow Statement

    #POPULATING PROFIT AND LOSS
    #B[x]:  x being odd gives the individual breakups of the different financial statements
    #A[3] has the children responsible for the financial statements
    A=list(table[2].children)
    B=list(A[3].children)

    PROFIT_LOSS["SALES"]=Year_wise_component_breakup(list(B[1].children))
    PROFIT_LOSS["EXPENSES"]=Year_wise_component_breakup(list(B[3].children))
    PROFIT_LOSS["OPERATING_PROFIT"]=Year_wise_component_breakup(list(B[5].children))
    PROFIT_LOSS["OPM_PER"]=Year_wise_component_breakup(list(B[7].children))
    PROFIT_LOSS["OTHER_INCOME"]=Year_wise_component_breakup(list(B[9].children))
    PROFIT_LOSS["INTEREST"]=Year_wise_component_breakup(list(B[11].children))
    PROFIT_LOSS["DEPRICIATION"]=Year_wise_component_breakup(list(B[13].children))
    PROFIT_LOSS["PBT"]=Year_wise_component_breakup(list(B[15].children))
    PROFIT_LOSS["TAX_PER"]=Year_wise_component_breakup(list(B[17].children))
    PROFIT_LOSS["NET_PROFIT"]=Year_wise_component_breakup(list(B[19].children))
    PROFIT_LOSS["EPS"]=Year_wise_component_breakup(list(B[21].children))
    PROFIT_LOSS["DIVIDEND_PAYOUT"]=Year_wise_component_breakup(list(B[23].children))
    print(PROFIT_LOSS["SALES"])
    #POPULATING BALANCE_SHEET
    #A[3] has the children responsible for the financial statements
    A=list(table[3].children)
    B=list(A[3].children)

    BALANCE_SHEET["SHARE_CAPITAL"]=Year_wise_component_breakup(list(B[1].children))
    BALANCE_SHEET["RESERVES"]=Year_wise_component_breakup(list(B[3].children))
    BALANCE_SHEET["BORROWINGS"]=Year_wise_component_breakup(list(B[5].children))
    BALANCE_SHEET["OTHER_LIABILITIES"]=Year_wise_component_breakup(list(B[7].children))
    BALANCE_SHEET["TOTAL_LIABILITIES"]=Year_wise_component_breakup(list(B[9].children))
    BALANCE_SHEET["FIXED_ASSETS"]=Year_wise_component_breakup(list(B[11].children))
    BALANCE_SHEET["CWIP"]=Year_wise_component_breakup(list(B[13].children))
    BALANCE_SHEET["INVESTMENTS"]=Year_wise_component_breakup(list(B[15].children))
    BALANCE_SHEET["OTHER_ASSETS"]=Year_wise_component_breakup(list(B[17].children))
    BALANCE_SHEET["TOTAL_ASSETS"]=Year_wise_component_breakup(list(B[19].children))

    #POPULATING CASHFLOWS
    #A[3] has the children responsible for the financial statements
    A=list(table[4].children)
    B=list(A[3].children)

    CASHFLOW["OPERATING_ACTIVITY"]=Year_wise_component_breakup(list(B[1].children))
    CASHFLOW["INVESTING_ACTIVITY"]=Year_wise_component_breakup(list(B[3].children))
    CASHFLOW["FINANCING_ACTIVITY"]=Year_wise_component_breakup(list(B[5].children))
    excel_write(PROFIT_LOSS,wb,sheet1,x)
    excel_write(BALANCE_SHEET,wb,sheet2,x)
    excel_write(CASHFLOW,wb,sheet3,x)
#------------------------------------------End of Writing to Excel---------------------------------------------------
#-----------------------------------------END of FUNCTIONS-------------------------------------------

if(__name__=="__main__"):
    a=str(sys.argv[1])
    populating_excel(a)
