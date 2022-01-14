import xlsxwriter
import requests
import datetime

# Pārmaina padoto šūnu uz nākamo šūnu, kurā būs jāievieto vērtība
def changeCell(cell):
    current_cell_letter = cell[0]
    current_cell_nr=(cell[1:])
    next_cell_nr_int = int(current_cell_nr)+1 
    new_cell = current_cell_letter+str(next_cell_nr_int)
    return new_cell

# Pievieno kolonnu virsrakstus un citu info 
workbook = xlsxwriter.Workbook('covid_data.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', "Dati ģenerēti izmantojot https://github.com/Ginters17/CovidDataToExcel. Datu avots: data.gov.lv" )
worksheet.write('A2', "Datums" )
worksheet.write('B2', "Testu skaits" )
worksheet.write('C2', "Covid-19 inficēto skaits" )
worksheet.write('D2', "Īpatsvars" )
worksheet.write('E2', "Inficēto skaits, kuri nav vakcinējušies" )
worksheet.write('F2', "Inficēto skaits, kuri ir vakcinējušies")
worksheet.write('G2', "Mirušo personu skaits")
worksheet.write('H2', "Vakcinēto mirušo personu skaits")
worksheet.write('I2', "Nevakcinēto vai vakcinācijas kursu nepabeigušo mirušo personu skaits")

# Iegūst JSON no data.gov.lv mājaslapas API
# Need to change offset value in URL after 100 records have been posted in JSON. Can write a script for this but i am too lazy, so just manually change it every 100 days lul.
response = requests.get('https://data.gov.lv/dati/lv/api/3/action/datastore_search?offset=600&resource_id=d499d2f0-b1ea-4ba2-9600-2c701b03bd4a') 
data = response.json()

# Izskaita cik ir kopā ierakstu par covid statistiku katru dienu, lai varētu izvēlēties pēdējo (visjaunāko)
count = 0
for item in data["result"]["records"]:
    count+=1

# Cik daudz ierakstus ierakstīt excelī (max: 100, ja ievada manuāli). Count -1 = skaits ar ierakstiem json.
ierakstu_daudzums = count - 1 

AX = "A3"
BX = "B3"
CX = "C3"
DX = "D3"
EX = "E3"
FX = "F3"
GX = "G3"
HX = "H3"
IX = "I3"

while ierakstu_daudzums>0:
    # count - 1 is the last record in json. count -2 is previous to last record in json. 
    count = count - 1

    # Gets the stats from json. Also formats it.
    tests_count = data["result"]["records"][count]["TestuSkaits"]
    cases_count = data["result"]["records"][count]["ApstiprinataCOVID19InfekcijaSkaits"]
    dead_count = data["result"]["records"][count]["MirusoPersonuSkaits"]
    dead_count_vaccinated = data["result"]["records"][count]["MirusoPersonuSkaits_Vakc"]
    dead_count_unvaccinated = data["result"]["records"][count]["MirusoPersonuSkaits_NevakcVakcNepab"]
    proportion = data["result"]["records"][count]["Ipatsvars"] 
    cases_count_unvaccinated = data["result"]["records"][count]["ApstCOVID19InfSk_Nevakc"] 
    cases_count_vaccinated = data["result"]["records"][count]["ApstCOVID19InfSk_Vakc"] 
    cumulitive_count = data["result"]["records"][count]["14DienuKumulativaSaslimstibaUz100000Iedzivotaju"] 
    date_and_time = data["result"]["records"][count]["Datums"]
    size = len(date_and_time)

    date = date_and_time[:size - 9] # Subtracts hour:minute:second part of string
    datetime_object = datetime.datetime.strptime(date, '%Y-%m-%d') # Turns date into datetime object
    Previous_Date_and_time = datetime_object - datetime.timedelta(days=1) # Subtracts 1 day from datetime object because date in json is 1 day further
    size = len(str(Previous_Date_and_time))
    Previous_Date = str(Previous_Date_and_time)[:size - 9] # Subtracts hour:minute:second part of string
    date = Previous_Date # Assign just calculated previous date to variable date for easier understanding


    worksheet.write(AX, date)
    worksheet.write(BX, tests_count)
    worksheet.write(CX, cases_count)
    worksheet.write(DX, proportion)
    worksheet.write(EX, cases_count_unvaccinated)
    worksheet.write(FX, cases_count_vaccinated)
    worksheet.write(GX, dead_count)
    worksheet.write(HX, dead_count_vaccinated)
    worksheet.write(IX, dead_count_unvaccinated)

    ierakstu_daudzums = ierakstu_daudzums - 1

    AX = changeCell(AX)
    BX = changeCell(BX)
    CX = changeCell(CX)
    DX = changeCell(DX)
    EX = changeCell(EX)
    FX = changeCell(FX)
    GX = changeCell(GX)
    HX = changeCell(HX)
    IX = changeCell(IX)

workbook.close()