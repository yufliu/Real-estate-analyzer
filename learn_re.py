# IT WORKS
# https://imgbb.com/
import pandas as pd
from selenium import webdriver
import time

import xlwt
import sys
import csv
import math
from mortgage import Loan
import decimal
import subprocess
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import plotly.express as px
#from openpyxl import load_workbook
# insert at 1, 0 is the script path (or '' in REPL)
sys.path.insert(0, '/Users/user/Documents/Atom/Python/MLH/Functions')
from getDigitfromString import getDigitfromString
from not_available import not_available
RE_data_path = "/Users/user/Documents/Atom/Python/MLH"
# ######### CHANGE HERE!########################
# basic info
ARV = 0
print("Hi and welcome! This will teach you the numbers behind Real Estate Investing. \nIt will begin by asking you some important questions. \nAnswer to the best of your abilities and if you do not know the answers, simply press 'ENTER' and it will input default values. \n")
DownPaymentPer = (input("How much money down (in %) are you putting (default is 20%): ") or \
       20)
Interest = (input("What's the interest you will have for your loan (default is 3.5%): ") or \
       3.5)

# unaccounted (10%)
CapexP = input("CapEx or 'Capital Expenditures' are big ticket items that you might need to pay for later such as fixing a roof, air conditioner, etc. This is usually 5% on average. \nWhat is your estimated CapEX % (Default is 5%): ") or 5
MaintP = input("For the times you need to hire general maintenance, plumber to fix the pipes or electrician for something. \nWhat is your maintenance budget in % (Default is 5%): ") or 5
VacancyP = input("You need to account when someone will not be living in your home. You can look for the vacancy rate in your city. \nWhat is your vacancy rate in % (Default is 5%): ") or 5
P_MgmtP = input("If you need someone to manage your property, put the rate at which they charge (usually 10%). If you don't need one, put 0 (Default is 0): ") or 0
# ######### END CHANGE HERE!###############
url = (input("Paste in the URL of the listing you want to get data for: ") or \
        "https://www.zillow.com/homedetails/3766-Laurel-Trace-Way-Tallahassee-FL-32303/80779460_zpid/")
       #"https://www.zillow.com/homedetails/2525-Hartsfield-Rd-APT-10-Tallahassee-FL-32303/58576263_zpid/")
#url = "https://www.zillow.com/homedetails/1201-Burnham-Dr-San-Jose-CA-95132/19773843_zpid/"
# Open the browser and URL
driver = webdriver.Firefox()
driver.get(url)
driver.maximize_window()
time.sleep(2)

#house_cards = driver.find_elements_by_class_name("list-card-info")

# Create the arrays to hold the values
address = []
price = []
bedrooms = []
baths = []
sq_ft = []
income = []
link = []

# Expenses
mortgage = []
tax = []
insurance = []
HOA = []

############################


# INCOME
try:
    #income_x = (driver.find_element_by_xpath('(//div[@class="Spacer-sc-17suqs2-0 pfWXf"])[2]/p').text)
    income_ws = driver.find_element_by_xpath('//span[contains(text(),"Rent Zestimate")]/following::p[1]')
    income_x = (driver.find_element_by_xpath('//span[contains(text(),"Rent Zestimate")]/following::p[1]').text)
    driver.execute_script('arguments[0].setAttribute("style", "color: blue; border: 2px solid orange; background:yellow")', income_ws)
except:
    print("error")
    income.append("Not Available")
else:
    print("income found")
    income.append(income_x)
time.sleep(2)
#print(income[house_cards[house]])

# EXPENSES
                                        #sc-Rmtcm GNsVX  sc-Rmtcm fzcNkq         Text-aiai24-0 cJeryq                        Text-aiai24-0 cJeryq
try:                                                         #sc-Rmtcm GNs
    #mortgage_x = (driver.find_element_by_xpath('.//div[@class="sc-Rmtcm GNsVX"]/span[@class="Text-aiai24-0 cJeryq"]').text)
    mortgage_ws = driver.find_element_by_xpath("//span[text()='Principal & interest']/following-sibling::span")
    mortgage_x = (driver.find_element_by_xpath("//span[text()='Principal & interest']/following-sibling::span").text)
    driver.execute_script('arguments[0].setAttribute("style", "color: blue; border: 2px solid orange; background:yellow")', mortgage_ws)
except:
    mortgage.append("N/A")
    #print("MESSED UP")
else:
    mortgage.append(mortgage_x)
    #print(mortgage_x)

try:
    #tax_x = (driver.find_element_by_xpath('(.//div[@class="sc-Rmtcm GNsVX"]/span[@class="Text-aiai24-0 cJeryq"])[3]').text)
    tax_ws = driver.find_element_by_xpath("//span[text()='Property taxes']/following-sibling::span")
    tax_x = (driver.find_element_by_xpath("//span[text()='Property taxes']/following-sibling::span").text)
    driver.execute_script('arguments[0].setAttribute("style", "color: blue; border: 2px solid orange; background:yellow")', tax_ws)
except:
    tax.append("N/A")
else:
    tax.append(tax_x)

try:
    #insurance_x = (driver.find_element_by_xpath('(.//div[@class="sc-Rmtcm GNsVX"]/span[@class="Text-aiai24-0 cJeryq"])[4]').text
    insurance_ws = driver.find_element_by_xpath("//span[text()='Home insurance']/following-sibling::span")
    insurance_x = (driver.find_element_by_xpath("//span[text()='Home insurance']/following-sibling::span").text)
    driver.execute_script('arguments[0].setAttribute("style", "color: blue; border: 2px solid orange; background:yellow")', insurance_ws)
except:
    insurance.append("N/A")
else:
    insurance.append(insurance_x)

try:
    #HOA_x = (driver.find_element_by_xpath('(.//div[@class="sc-Rmtcm GNsVX"]/span[@class="Text-aiai24-0 cJeryq"])[5]').text)
    HOA_ws = driver.find_element_by_xpath("//span[text()='HOA fees']/following-sibling::span")
    HOA_x = (driver.find_element_by_xpath("//span[text()='HOA fees']/following-sibling::span").text)
    driver.execute_script('arguments[0].setAttribute("style", "color: blue; border: 2px solid orange; background:yellow")', HOA_ws)

except:
    HOA.append("N/A")
else:
    HOA.append(HOA_x)

link.append(driver.current_url)

# Close Listing
time.sleep(1)


# Get the basic house information
address.append(driver.find_element_by_xpath('.//h1[@class="ds-address-container"]/span').text)
address_ws = driver.find_element_by_xpath('.//h1[@class="ds-address-container"]/span')
driver.execute_script('arguments[0].setAttribute("style", "color: blue; border: 2px solid orange; background:yellow")', address_ws)       #Highlighting stuff on page
#driver.execute_script("arguments[0].setAttribute('style','border: 4px solid red');", address0)
#driver.execute_script("arguments[0].setAttribute('style', arguments[1]);",address0, "background:yellow")
#driver.execute_script("arguments[0].setAttribute('style', arguments[1]);", address0, "background:yellow; color: Red")
price_ws = driver.find_element_by_class_name("ds-value")
price.append(price_ws.text)
driver.execute_script('arguments[0].setAttribute("style", "color: blue; border: 2px solid orange; background:yellow")', price_ws)
bedrooms_ws = driver.find_element_by_xpath('(//div[@class="ds-chip"]//span[@class="ds-bed-bath-living-area"]/span[ not(@class="ds-summary-row-label-secondary")])[1]')
bedrooms.append(driver.find_element_by_xpath('(//div[@class="ds-chip"]//span[@class="ds-bed-bath-living-area"]/span[ not(@class="ds-summary-row-label-secondary")])[1]').text)
driver.execute_script('arguments[0].setAttribute("style", "color: blue; border: 2px solid orange; background:yellow")', bedrooms_ws)
#baths.append(driver.find_element_by_xpath('.//button[@class="TriggerText-sc-139r5uq-0 jfjsxZ TooltipPopper-io290n-0 sc-jlyJG eVrWvb"]/span[@class="ds-bed-bath-living-area"][1]').text)
baths_ws = driver.find_element_by_xpath('(//div[@class="ds-chip"]//span[@class="ds-bed-bath-living-area"]/span[ not(@class="ds-summary-row-label-secondary")])[2]')
baths.append(driver.find_element_by_xpath('(//div[@class="ds-chip"]//span[@class="ds-bed-bath-living-area"]/span[ not(@class="ds-summary-row-label-secondary")])[2]').text)
driver.execute_script('arguments[0].setAttribute("style", "color: blue; border: 2px solid orange; background:yellow")', baths_ws)
print("bath: " + str(baths))

#sq_ft.append(driver.find_element_by_xpath('.//h3[@class="ds-bed-bath-living-area-container"]/span[@class="ds-bed-bath-living-area"][3]').text)
sq_ft_ws = driver.find_element_by_xpath('(//div[@class="ds-chip"]//span[@class="ds-bed-bath-living-area"]/span[ not(@class="ds-summary-row-label-secondary")])[3]')
sq_ft.append(driver.find_element_by_xpath('(//div[@class="ds-chip"]//span[@class="ds-bed-bath-living-area"]/span[ not(@class="ds-summary-row-label-secondary")])[3]').text)
driver.execute_script('arguments[0].setAttribute("style", "color: blue; border: 2px solid orange; background:yellow")', sq_ft_ws)
print("sqft:" + str(sq_ft))

#driver.quit()


#print(address, bedrooms, bathrooms, sq_ft, price, income, mortgage, tax, insurance, HOA)
listings = pd.DataFrame({
                                    "address":address,
                                    "bedrooms":bedrooms,
                                    "baths":baths,
                                    "sq_ft":sq_ft,
                                    "price":price,
                                    "income":income,
                                    "mortgage":mortgage,
                                    "tax":tax,
                                    "insurance":insurance,
                                    "HOA":HOA,
                                    "link":link
                                    },)

#print(listings)

csvfile = "one_house_data.csv"
listings.to_csv(csvfile, sep='\t', encoding='utf-8', index=False, header=True)
print("Finished! View your csv file: " + csvfile)

####################################################################
##############################START ANALYSIS########################
####################################################################

file = csvfile
df = pd.read_csv(file, delimiter='\t')  # store data inside df
count = 0  # counter
#subprocess.call(['open', csvfile])

# New columns to add to file
#print(len(df['address']))
Total_Rent_M_c = [0] * len(df['address'])
Total_Rent_Y_c = [0] * len(df['address'])
Total_Exp_M_c = [0] * len(df['address'])
Total_Exp_Y_c = [0] * len(df['address'])
CashFlow_cm = [0] * len(df['address'])
CashFlow_cy = [0] * len(df['address'])
Net_Margin_c = [0] * len(df['address'])
CoCr_c = [0] * len(df['address'])
Cap_Rate_c = [0] * len(df['address'])
Good_c = [0] * len(df['address'])

########## START FOR LOOP ###################
for address in df['address']:
    #print(str(count) + ": " + address)
    #print(str(df['bedrooms'][count]) + " " + str(df['baths'][count]) + " | " + str(df['sq_ft'][count]))
    print("Price: " + (df['price'][count]))
    Price = getDigitfromString(df['price'][count])

    if Price == 0:
        print("Data Unavailable for listing #" + str(count))
        not_available(count, len(df['address']))
        count = count + 1
        print()
        continue

    # print("Income: " + (df['income'][count]))
    Rent_M = getDigitfromString(df['income'][count])

    if Rent_M == 0:
        print("Data Unavailable for listing #" + str(count))
        not_available(count, len(df['address']))
        count = count + 1
        print()
        continue

    # print("Tax: " + (df['tax'][count]))
    TaxesM = getDigitfromString(str(df['tax'][count]))
    # print("Insurance: " + (df['insurance'][count]))
    InsuranceM = getDigitfromString(str(df['insurance'][count]))
    # print("HOA: " + str((df['HOA'][count])))
    HOA_M = getDigitfromString(str(df['HOA'][count]))

    Rent_Y = int(Rent_M) * 12
    Rent_Y = float(Rent_Y)
    DownPayment_amt = float(DownPaymentPer)/100 * float(Price)
    Total_down = DownPayment_amt + float(ARV)
    Total_House_Cost = float(Price) + float(ARV)  # total price of house including ARV
    Capex_amtM = float(CapexP)/100 * float(Rent_M)
    Maint_amtM = float(MaintP)/100 * float(Rent_M)
    Vacancy_amtM = float(VacancyP)/100 * float(Rent_M)
    P_Mgt_amtM = float(P_MgmtP)/100 * float(Rent_M)

#-------Mortgage block---------
    loan = Loan(principal=(float(Price) - DownPayment_amt), interest=float(Interest)/100, term=30)
    Mortgage_M = float(loan.monthly_payment)
    # print("Mortgage: " + (df['mortgage'][count]) + " VS " + str(Mortgage_M))
    Mortgage = getDigitfromString(str(df['mortgage'][count]))

    if Mortgage >= Mortgage_M:
        Mortgage_M = Mortgage
    else:
        Mortage_M = Mortgage_M
#--------------------------------
    # Summary Metrics
    Total_Exp_M = float(float(TaxesM) + float(InsuranceM) + Capex_amtM + Maint_amtM + Vacancy_amtM + P_Mgt_amtM + Mortgage_M + HOA_M)
    Total_Exp_Y = Total_Exp_M * 12
    NOI_M = decimal.Decimal(int(Rent_M) - decimal.Decimal(Total_Exp_M) + decimal.Decimal(Mortgage_M))  # Income - Expenses (excluding mortgage)
    NOI_Y = float(NOI_M) * 12

    CashFlowM = float(Rent_M) - Total_Exp_M
    CashFlowY = CashFlowM * 12
    Net_Margin = (CashFlowM / float(Rent_M)) * 100
    CoCr = (CashFlowY / Total_down) * 100  # Verify this
    Cap_Rate = (NOI_Y / Total_House_Cost) * 100

    # Put into column arrays
    Total_Rent_M_c[count] = (("%.2f" % int(Rent_M)))
    Total_Rent_Y_c[count] = (("%.2f" % Rent_Y))
    Total_Exp_M_c[count] = (("%.2f" % Total_Exp_M))
    Total_Exp_Y_c[count] = (("%.2f" % Total_Exp_Y))
    CashFlow_cm[count] = (("%.2f" % CashFlowM))
    CashFlow_cy[count] = (("%.2f" % CashFlowY))

    Net_Margin_c[count] = ("%.3f" % Net_Margin + "%")
    CoCr_c[count] = (("%.3f" % CoCr + "%"))
    Cap_Rate_c[count] = (("%.3f" % Cap_Rate + "%"))

    #------CREATE PIE CHART
    def func(pct, allvals):
        absolute = int(pct/100.*np.sum(allvals))
        return "{:.2f}%\n(${:d})".format(pct, absolute)

    labels = 'Taxes', 'Insurance', 'Capex', 'Maintenance', 'vacancy', 'Mgmt', 'Mortgage', 'HOA'
    sizes = [float(TaxesM), float(InsuranceM), Capex_amtM, Maint_amtM, Vacancy_amtM, P_Mgt_amtM, Mortgage_M, HOA_M]

    explode = (0, 0, 0, 0, 0, 0, 0.1, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')

    fig1, ax1 = plt.subplots()
    ax1.pie(sizes, explode=explode, labels=labels, autopct=lambda pct: func(pct, sizes),
            shadow=True, startangle=90)
    ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    #ax1.set_title("Monthly Expenses: $" + str(Total_Exp_M))
    plt.title("Monthly Expenses: $" + str(Total_Exp_M))
    plt.savefig("fig.png")
    # debug
    # print(len(Total_Rent_M_c))
    # print(len(Total_Rent_Y_c))
    # print(len(Total_Exp_M_c))
    # print(len(Total_Exp_Y_c))
    # print(len(CashFlow_cm))
    # print(len(CashFlow_cy))
    # print(len(CoCr_c))
    # print(len(Cap_Rate_c))
    # print((count))

    if CoCr > 5:
        print("NOTE: CoCR PASSSES")
        Good_c[count] = ("X")
    else:
        print("NOTE: CoCR DOES NOT PASS MIN 5%")
        Good_c[count] = (" ")
    print("---------------------END---------------------")

    print("")
    count = count + 1

# df['RENT /m'] = Total_Rent_M_c
# df['RENT /y'] = Total_Rent_Y_c
# df['EXPENSE /m'] = Total_Exp_M_c
# df['EXPENSES /y'] = Total_Exp_Y_c
# df['*Cash Flow /m*'] = CashFlow_cm
# df['*Cash Flow /y*'] = CashFlow_cy
# df['Net Margin'] = Net_Margin_c
# df['CoCr'] = CoCr_c
# df['Cap Rate'] = Cap_Rate_c
# df['Pursue'] = Good_c

# # Changing the order of the columns; put links as the last col
# cols = list(df.columns.values)
# df = df[cols[0:9] + cols[11:22] + [cols[10]]]
# #df = df[cols[0:8]] + cols[10:] + [cols[-1]]
#
# # Sort the values
# try:
#     df = df.sort_values('*Cash Flow /m*', ascending=False)
# except:
#     print("THERE WAS A PROBLEM!")
# else:
#     print("Nothing went wrong")
# df = df.sort_values(['price', '*Cash Flow /m*'], ascending=[1, 1])

# CSV
# outfile = "one_house.csv"
# df.to_csv(outfile, sep='\t', encoding='utf-8', index=False, header=True)
# print("Finished! View your csv file: " + outfile)

#-----------------START MAKING IT LOOK NICE--------------------
#---------------------------------------------------------------
df1 = pd.DataFrame({
    'Address': [address],
    'Bedrooms': [bedrooms[0]],
    'Bathrooms': [baths[0]],
    'Sq. Ft.': [sq_ft[0]],
    'Link': [link[0]],
    })

df_inc = pd.DataFrame({
    'Income': ['rent'],
    'Month': float(Total_Rent_M_c[0]),
    #'Month': 33000.00,
    'Year': float(Total_Rent_Y_c[0])
    })

df_exp = pd.DataFrame({
    'Expenses': ['Mortgage', 'Taxes', 'Insurance', 'HOA', 'Utilities', 'Capex'+' ('+str(CapexP)+'%)', 'Repairs'+' ('+str(MaintP)+'%)', 'Vacancy'+' ('+str(VacancyP)+'%)', 'Management'+' ('+str(P_MgmtP)+'%)'],
    'Month': [Mortgage_M,TaxesM,InsuranceM,HOA_M,'',Capex_amtM,Maint_amtM,Vacancy_amtM,P_Mgt_amtM],
    'Year': [Mortgage_M*12,TaxesM*12,InsuranceM*12,HOA_M*12,'-',Capex_amtM*12,Maint_amtM*12,Vacancy_amtM*12,P_Mgt_amtM*12],
    })

df_details = pd.DataFrame({
    'Metrics': ['Income', 'Expenses', 'Cash Flow', 'COCR', 'Cap Rate', 'NOI', 'Net Margin', 'Pursue'],
    # 'Month':   ['$'+str(Total_Rent_M_c[0]),   Total_Exp_M_c[0], CashFlow_cm[0] ,'-','-','$'+str(NOI_M), '-', ''],
    # 'Year': ['$'+str(Total_Rent_Y_c[0]),   Total_Exp_Y_c[0], CashFlow_cy[0], CoCr_c[0], Cap_Rate_c[0], '$'+str(NOI_Y), Net_Margin_c[0], Good_c[0]]
    'Month':   [float(Total_Rent_M_c[0]),   float(Total_Exp_M_c[0]), float(CashFlow_cm[0]),'-','-',float(NOI_M), '-', ''],
    'Year':    [float(Total_Rent_Y_c[0]),   float(Total_Exp_Y_c[0]), float(CashFlow_cy[0]), CoCr_c[0], Cap_Rate_c[0], float(NOI_Y), Net_Margin_c[0], Good_c[0]]
    })

Price = getDigitfromString(price[0])
Closing_per = 2
closing = Price*Closing_per/100
Total_down = closing + Total_down
Loan_amount = Price-DownPayment_amt

df_prop = pd.DataFrame({
    'Property Details': ['Price Listed', 'ARV', 'Closing (' + str(Closing_per) + "%)", 'Down Payment (' + str(DownPaymentPer) + "%)", 'Finance needed', 'Int Rate', 'Total $$ Needed'],
    #'$': [price, ARV, str(int(price)*.02), DownPayment_amt, '', Interest, str(int(price)*.02+ARV)],
    # '$': ['$'+str(price[0]), '$'+str(ARV), '$'+str(closing), '$'+str(DownPayment_amt), '$'+str(Loan_amount), str(Interest)+'%', '$'+str(Total_down)],
    'Amount': [Price, float(ARV), float(closing), float(DownPayment_amt), float(Loan_amount), str(Interest)+'%', float(Total_down)],
    })

# Excel
outfile = RE_data_path + "one_house_complete.xlsx"
#df.to_excel(outfile, encoding='none', index=False, header=True)
# print("Finished! View your csv file: " + outfile)
FileName = str(outfile)
writer = pd.ExcelWriter(FileName, engine='xlsxwriter')

# Add each section to Excel sheet

df_inc.to_excel(writer, startrow=1, startcol=1, index=False)
df_exp.to_excel(writer, startrow=4, startcol=1, index=False)
df_details.to_excel(writer, startrow=1, startcol=5, index=False)
df1.to_excel(writer, startrow=1, startcol=9, index=False)
df_prop.to_excel(writer, sheet_name='Sheet1', startrow=4, startcol=9, header=True, index=False)

workbook  = writer.book
worksheet = writer.sheets['Sheet1']
print("Finished! View your excel file: " + outfile)

# FORMATS-----------

# Add a header format.
income_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#A0CD63',
    'underline': True,
    'num_format': '$####,##0',
    'border': 1})

COCR_format = workbook.add_format({'bold': True, 'fg_color': '#CEEDD0', 'border': 1, 'num_format': '0.00%', 'align': 'right'})
COCR_bad_format = workbook.add_format({'bold': True, 'fg_color': '#F7C9CF', 'border': 1, 'num_format': '0.00%', 'align': 'right'})
expense_format = workbook.add_format({'bold': True, 'fg_color': '#EF857D', 'border': 1})
details_format = workbook.add_format({'bold': True, 'fg_color': '#4BAFEA', 'border': 1})
header_format = workbook.add_format({'bold': True, 'fg_color': '#EBE9DD', 'border': 1})
money = workbook.add_format({'num_format': '$#,##0', 'align': 'right'})
percent = workbook.add_format({'num_format': '0.00%', 'align': 'right'})



# Add header format for each section
for col_num, value in enumerate(df_inc.columns.values):
    worksheet.write(1, col_num + 1, value, income_format)

for row, value in enumerate(df_inc.Month.values):
    worksheet.write(row + 2, 2, value, money)

for row, value in enumerate(df_inc.Year.values):
    worksheet.write(row + 2, 3, value, money)
# ---
for col_num, value in enumerate(df_exp.columns.values):
    worksheet.write(4, col_num + 1, value, expense_format)

for row, value in enumerate(df_exp.Month.values):
    worksheet.write(row + 5, 2, value, money)

for row, value in enumerate(df_exp.Year.values):
    worksheet.write(row + 5, 3, value, money)
# ---
for col_num, value in enumerate(df_details.columns.values):
    worksheet.write(1, col_num + 5, value, details_format)

for row, value in enumerate(df_details.Month.values):
    worksheet.write(row + 2, 6, value, money)

for row, value in enumerate(df_details.Year.values):
    worksheet.write(row + 2, 7, value, money)
# ---
for col_num, value in enumerate(df1.columns.values):
    worksheet.write(1, col_num + 9, value, header_format)

for col_num, value in enumerate(df_prop.columns.values):
    worksheet.write(4, col_num + 9, value, header_format)

for row, value in enumerate(df_prop.Amount.values):
    worksheet.write(row + 5, 10, value, money)


# worksheet.set_column(1, 11, 15)
worksheet.set_column('B:B', 15)
worksheet.set_column('F:F', 15)
worksheet.set_column('J:J', 17)
worksheet.set_column('K:K', 10)
worksheet.set_column('N:N', 50)

# Dynamics:
worksheet.write_formula('D3', '=sum(C3*12)', money)
worksheet.write_formula('D6', '=C6*12', money)
worksheet.write_formula('D7', '=C7*12', money)
worksheet.write_formula('D8', '=C8*12', money)
worksheet.write_formula('D9', '=C9*12', money)
worksheet.write_formula('D10', '=C10*12', money)
worksheet.write_formula('D11', '=C11*12', money)
worksheet.write_formula('D12', '=C12*12', money)
worksheet.write_formula('D13', '=C13*12', money)
worksheet.write_formula('D14', '=C14*12', money)

worksheet.write_formula('G3', '=C3', money)     # Income M
worksheet.write_formula('H3', 'D3', money)
worksheet.write_formula('G4', '=sum(C6:C14)', money)    # expense M
worksheet.write_formula('H4', '=sum(D6:D14)', money)
worksheet.write_formula('G5', '=G3-G4', money)       # Cash Flow M
worksheet.write_formula('H5', '=H3-H4', money)
worksheet.write_formula('H6', '=H5/K12', percent)      # COCR
worksheet.write_formula('H7', '=H8/(K6+K7)', percent)      # Cap rate
worksheet.write_formula('G8', '=G3-(G4-C6)', money)      # NOI M
worksheet.write_formula('H8', '=H3-(H4-D6)', money)      # NOI
worksheet.write_formula('H9', '=G5/G3', percent)

# Other stuff
worksheet.write('N15', "test text")

if (CoCr >= 5):
    worksheet.write('H6', '=H5/K12', COCR_format)
else:
    worksheet.write('H6', '=H5/K12', COCR_bad_format)

writer.save()

# Pie chart using plotly----------------------------------------
title_pie = "Monthly Expenses: $" + str(Total_Exp_M)
sizes = [float(TaxesM), float(InsuranceM), Capex_amtM, Maint_amtM, Vacancy_amtM, P_Mgt_amtM, Mortgage_M, HOA_M]
fig = px.pie(df_exp, values='Month', names='Expenses', title=title_pie)
fig.update_traces(textposition='inside', textinfo='percent+label+value')
#fig.show()  # If you want to see it on the browser

fig.to_image(format="jpeg", engine="kaleido")
img_bytes = fig.to_image(format="png", width=600, height=350, scale=2)
fig.write_image("fig1.png")

#----Image
wb = openpyxl.load_workbook(FileName)
ws = wb.active
#img = openpyxl.drawing.image.Image('fig.png')
img = openpyxl.drawing.image.Image('fig1.png')
ws.add_image(img, 'B17')     # pie chart placement
wb.save(FileName)

print("Opening Excel File now")
subprocess.call(['open', FileName])
subprocess.call(['open', "/Users/user/Downloads/fullpage_js/index.html"])
#plt.show()
