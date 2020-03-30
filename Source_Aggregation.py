#Data Source: https://github.com/CSSEGISandData/COVID-19
#COVID-19 Aggregation Script by Cyber City Circuits

import csv, os
import openpyxl
from openpyxl import Workbook
import datetime

wb_name = '-Aggregated.xlsx'

cwd_dir = os.getcwd()
data_dir = os.getcwd()+"\\data\\"

death = 0
infected = 0
line = 1
line_total = 1

workbook = Workbook()
ws = workbook.active

sheet_co     = workbook.create_sheet('CO Totals', 0)
sheet_tx     = workbook.create_sheet('TX Totals', 0)
sheet_ca     = workbook.create_sheet('CA Totals', 0)
sheet_fl     = workbook.create_sheet('FL Totals', 0)
sheet_wa     = workbook.create_sheet('WA Totals', 0)
sheet_nc     = workbook.create_sheet('NC Totals', 0)
sheet_ny     = workbook.create_sheet('NY Totals', 0)
sheet_sc     = workbook.create_sheet('SC Totals', 0)
sheet_ga     = workbook.create_sheet('GA Totals', 0)
sheet_us     = workbook.create_sheet('US Totals', 0)
sheet_totals = workbook.create_sheet('Projections', 0)

date_object = datetime.date.today()

sheet_totals['A'+str(line_total)] = 'US Death Count Projections'
line_total += 1

def analysis_state(sheet_name, state_name):
    global line, line_total
    print("Parsing State: "+state_name)
    death = 0
    infected = 0
    line = 1
    sheet_name['A'+str(line)] = (state_name)
    line += 1
    sheet_name['A'+str(line)] = ('Date')
    sheet_name['B'+str(line)] = ('Infected')
    sheet_name['C'+str(line)] = ('Rate')
    sheet_name['D'+str(line)] = ('Deaths')
    sheet_name['E'+str(line)] = ('Rate')
    sheet_name['F'+str(line)] = ('Avaerage (7-Day)')
    sheet_name.column_dimensions['A'].width = 10
    sheet_name.column_dimensions['F'].width = 16
    line += 1 
    for fname in os.listdir(data_dir):
        death = 0
        infected = 0
        recovered = 0
        average_7 = 0
        if fname.endswith(".csv"):
            date = fname.split('.')[0]
            with open(data_dir + fname) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=',')
                line_count = 0
                for row in csv_reader:
                    if row[3] == 'US':
                        if state_name == 'all':
                            death = death + int(row[8])
                            infected = infected + int(row[7])
                            recovered = recovered + int(row[9])
                        elif row[2].lower() == state_name:
                            death = death + int(row[8])
                            infected = infected + int(row[7])
                            recovered = recovered + int(row[9])
                    if row[1].lower() == 'us':
                        if state_name == 'all':
                            if row[4] != '':
                                death = death + int(row[4])
                            if row[3] != '':
                                infected = infected + int(row[3])
                            if row[5] != '':
                                recovered = recovered + int(row[5])
                        elif row[0].lower() == state_name:
                            if row[4] != '':
                                death = death + int(row[4])
                            if row[3] != '':
                                infected = infected + int(row[3])
                            if row[5] != '':
                                recovered = recovered + int(row[5])
                    line_count += 1
            sheet_name['A'+str(line)] = date
            sheet_name['B'+str(line)] = infected
            sheet_name['D'+str(line)] = death
            if sheet_name['B' + str(line-1)].value == 0:
                sheet_name['C' + str(line)] = '0'
                sheet_name['C'+str(line)].number_format = '00.00%'
            elif line == 3:
                sheet_name['C' + str(line)] = '0'
                sheet_name['C'+str(line)].number_format = '00.00%'
            else:
                sheet_name['C'+str(line)] = ('=(B' + str(line) + '/B' + str(line-1) + ') - 1')
                sheet_name['C'+str(line)].number_format = '00.00%'
            
            if sheet_name['D' + str(line-1)].value == 0:
                sheet_name['E' + str(line)] = '0'
                sheet_name['C'+str(line)].number_format = '00.00%'
            elif line == 3:
                sheet_name['E' + str(line)] = '0'
                sheet_name['C'+str(line)].number_format = '00.00%'
            else:
                sheet_name['E'+str(line)] = ('=(D' + str(line) + '/D' + str(line-1) + ') - 1')
                sheet_name['E'+str(line)].number_format = '00.00%'
            if sheet_name['E'+str(line)].value != '0':
                sheet_name['F'+str(line)] = ('=AVERAGE(E' + str(line-6) + ':E' + str(line)+')')
                average_7 = sheet_name['F'+str(line)].value
                sheet_name['F'+str(line)].number_format = '00.00%'
            else:
                sheet_name['F'+str(line)] = 0
                sheet_name['F'+str(line)].number_format = '00.00%'
            print ('Date: ' + date + ' Infected: ' + str(infected).rjust(7)
                + ' Death: ' + str(death).rjust(5)
                )
            line += 1
    line += 1
    sheet_name['A'+str(line)] = 'Average (7-Day):'
    sheet_name['C'+str(line)] = ('=AVERAGE(C' + str(line-8) + ':C' + str(line-2)+')')
    sheet_name['C'+str(line)].number_format = '00.00%'
    sheet_name['E'+str(line)] = ('=AVERAGE(E' + str(line-8) + ':E' + str(line-2)+')')
    sheet_name['E'+str(line)].number_format = '00.00%'
    #Projections
    line += 2
    sheet_name['A'+str(line)] = ('Projections')
    line += 1
    if (((sheet_name['D'+str(line-5)].value)  != 0) and 
        ((sheet_name['D'+str(line-6)].value)  != 0) and 
        ((sheet_name['D'+str(line-7)].value)  != 0) and 
        ((sheet_name['D'+str(line-8)].value)  != 0) and 
        ((sheet_name['D'+str(line-9)].value)  != 0) and 
        ((sheet_name['D'+str(line-10)].value) != 0) and 
        ((sheet_name['D'+str(line-11)].value) != 0)):
        average_7 = ((int(sheet_name['D'+str(line-5)].value)/ int(sheet_name['D'+str(line-6)].value)) + 
                    (int(sheet_name['D'+str(line-6)].value)/ int(sheet_name['D'+str(line-7)].value)) + 
                    (int(sheet_name['D'+str(line-7)].value)/ int(sheet_name['D'+str(line-8)].value)) + 
                    (int(sheet_name['D'+str(line-8)].value)/ int(sheet_name['D'+str(line-9)].value)) + 
                    (int(sheet_name['D'+str(line-9)].value)/ int(sheet_name['D'+str(line-10)].value)) + 
                    (int(sheet_name['D'+str(line-10)].value)/int(sheet_name['D'+str(line-11)].value)) + 
                    (int(sheet_name['D'+str(line-11)].value)/int(sheet_name['D'+str(line-12)].value))) / 7
    print()
    print()
    starting_death = int(sheet_name['D'+str(line-5)].value)
    starting_date = sheet_name['A'+str(line-5)].value
    #month, day, year = (int(x) for x in starting_date.split('-'))
    #print(type(month))
    #print(str(month) + '-' + str(day) + '-' + str(year))
    #day += 1
    #date = datetime.date(month, day, year)
    #print(date)
    print('Final Death Count (US): ' + str(starting_death))
    print('Seven Day Average Increase (US): ' + str(average_7))
    sheet_name['A'+str(line)] = ('Day +1')
    sheet_name['D'+str(line)] = (starting_death * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1
    sheet_name['A'+str(line)] = ('Day +2')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1
    sheet_name['A'+str(line)] = ('Day +3')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1
    sheet_name['A'+str(line)] = ('Day +4')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1
    sheet_name['A'+str(line)] = ('Day +5')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1
    sheet_name['A'+str(line)] = ('Day +6')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1
    sheet_name['A'+str(line)] = ('Day +7')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'

    #Add to totals sheet
    sheet_totals['A'+str(line_total)] = state_name.title()
    sheet_totals['B'+str(line_total)] = 'Projections'
    sheet_totals['C'+str(line_total)] = 'Increase'
    line_total += 1
    starting_death
    sheet_totals['A'+str(line_total)] = ("Today's Count")
    sheet_totals['B'+str(line_total)] = (starting_death)

    line_total += 1
    sheet_totals['A'+str(line_total)] = ('Day +1')
    sheet_totals['B'+str(line_total)] = (starting_death * average_7)
    sheet_totals['C'+str(line_total)] = (((int(sheet_totals['B'+str(line_total)].value)) / starting_death)-1)
    sheet_totals['B'+str(line_total)].number_format = '0'
    sheet_totals['C'+str(line_total)].number_format = '00.00%'
    line_total += 1
    sheet_totals['A'+str(line_total)] = ('Day +2')
    sheet_totals['B'+str(line_total)] = (sheet_totals['B'+str(line_total-1)].value * average_7)
    sheet_totals['C'+str(line_total)] = (((int(sheet_totals['B'+str(line_total)].value)) / starting_death)-1)
    sheet_totals['B'+str(line_total)].number_format = '0'
    sheet_totals['C'+str(line_total)].number_format = '00.00%'
    line_total += 1
    sheet_totals['A'+str(line_total)] = ('Day +3')
    sheet_totals['B'+str(line_total)] = (sheet_totals['B'+str(line_total-1)].value * average_7)
    sheet_totals['C'+str(line_total)] = (((int(sheet_totals['B'+str(line_total)].value)) / starting_death)-1)
    sheet_totals['B'+str(line_total)].number_format = '0'
    sheet_totals['C'+str(line_total)].number_format = '00.00%'
    line_total += 1
    sheet_totals['A'+str(line_total)] = ('Day +4')
    sheet_totals['B'+str(line_total)] = (sheet_totals['B'+str(line_total-1)].value * average_7)
    sheet_totals['C'+str(line_total)] = (((int(sheet_totals['B'+str(line_total)].value)) / starting_death)-1)
    sheet_totals['B'+str(line_total)].number_format = '0'
    sheet_totals['C'+str(line_total)].number_format = '00.00%'
    line_total += 1
    sheet_totals['A'+str(line_total)] = ('Day +5')
    sheet_totals['B'+str(line_total)] = (sheet_totals['B'+str(line_total-1)].value * average_7)
    sheet_totals['C'+str(line_total)] = (((int(sheet_totals['B'+str(line_total)].value)) / starting_death)-1)
    sheet_totals['B'+str(line_total)].number_format = '0'
    sheet_totals['C'+str(line_total)].number_format = '00.00%'
    line_total += 1
    sheet_totals['A'+str(line_total)] = ('Day +6')
    sheet_totals['B'+str(line_total)] = (sheet_totals['B'+str(line_total-1)].value * average_7)
    sheet_totals['C'+str(line_total)] = (((int(sheet_totals['B'+str(line_total)].value)) / starting_death)-1)
    sheet_totals['B'+str(line_total)].number_format = '0'
    sheet_totals['C'+str(line_total)].number_format = '00.00%'
    line_total += 1
    sheet_totals['A'+str(line_total)] = ('Day +7')
    sheet_totals['B'+str(line_total)] = (sheet_totals['B'+str(line_total-1)].value * average_7)
    sheet_totals['C'+str(line_total)] = (((int(sheet_totals['B'+str(line_total)].value)) / starting_death)-1)
    sheet_totals['B'+str(line_total)].number_format = '0'
    sheet_totals['C'+str(line_total)].number_format = '00.00%'
    line_total += 1
    line_total += 1
    print()
    print()

def collect_totals():
    list_sheets = workbook.sheetnames
    for i in range(len(list_sheets)):
        print(list_sheets[i])

analysis_state(sheet_us, 'all')
analysis_state(sheet_ga, 'georgia')
analysis_state(sheet_sc, 'south carolina')
analysis_state(sheet_nc, 'north carolina')
analysis_state(sheet_ny, 'new york')
analysis_state(sheet_wa, 'washington')
analysis_state(sheet_ca, 'california')
analysis_state(sheet_fl, 'florida')
analysis_state(sheet_co, 'colorado')
analysis_state(sheet_tx, 'texas')

sheet_totals.column_dimensions['A'].width = 13
sheet_totals.column_dimensions['B'].width = 10
sheet_totals.column_dimensions['C'].width = 10
sheet_totals.column_dimensions['B'].alignment = Alignment(horizontal='right')
sheet_totals.column_dimensions['C'].alignment = Alignment(horizontal='right')

workbook.save(filename=cwd_dir + '\\' + str(date_object) + wb_name)

print()
print("Complete")
print()
print()


