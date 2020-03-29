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

workbook = Workbook()
ws = workbook.active


sheet_ca = workbook.create_sheet('CA Totals', 0)
sheet_wa = workbook.create_sheet('WA Totals', 0)
sheet_nc = workbook.create_sheet('NC Totals', 0)
sheet_ny = workbook.create_sheet('NY Totals', 0)
sheet_sc = workbook.create_sheet('SC Totals', 0)
sheet_ga = workbook.create_sheet('GA Totals', 0)
sheet_us = workbook.create_sheet('US Totals', 0)


def analysis_us():
    death = 0
    infected = 0
    line = 1

    sheet_us['A'+str(line)] = ('Date')
    sheet_us['B'+str(line)] = ('Infected')
    sheet_us['C'+str(line)] = ('Rate')
    sheet_us['D'+str(line)] = ('Deaths')
    sheet_us['E'+str(line)] = ('Rate')
    #sheet_us['F'+str(line)] = ('Recovered')
    #sheet_us['G'+str(line)] = ('Rate')
    sheet_us['F'+str(line)] = ('Avaerage (7-Day)')

    sheet_us.column_dimensions['A'].width = 10
    sheet_us.column_dimensions['F'].width = 16

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
                        death = death + int(row[8])
                        infected = infected + int(row[7])
                        recovered = recovered + int(row[9])
                    if row[1] == 'US':
                        if row[4] != '':
                            death = death + int(row[4])
                        if row[3] != '':
                            infected = infected + int(row[3])
                        if row[5] != '':
                            recovered = recovered + int(row[5])

                    line_count += 1
                #print(f'Processed {line_count} lines.')    

            sheet_us['A'+str(line)] = date
            sheet_us['B'+str(line)] = infected
            sheet_us['D'+str(line)] = death
            #sheet_us['F'+str(line)] = recovered

            #print(type(sheet_us['B'+str(line)]))

            if sheet_us['B' + str(line-1)].value == 0:
                sheet_us['C' + str(line)] = 'UNDEF'
            elif line == 2:
                sheet_us['C' + str(line)] = 'UNDEF'
            else:
                sheet_us['C'+str(line)] = ('=(B' + str(line) + '/B' + str(line-1) + ') - 1')
                sheet_us['C'+str(line)].number_format = '00.00%'
            
            if sheet_us['D' + str(line-1)].value == 0:
                sheet_us['E' + str(line)] = 'UNDEF'
            elif line == 2:
                sheet_us['E' + str(line)] = 'UNDEF'
            else:
                sheet_us['E'+str(line)] = ('=(D' + str(line) + '/D' + str(line-1) + ') - 1')
                sheet_us['E'+str(line)].number_format = '00.00%'

            #if sheet_us['F' + str(line-1)].value == 0:
            #    sheet_us['G' + str(line)] = 'UNDEF'
            #elif line == 2:
            #    sheet_us['G' + str(line)] = 'UNDEF'
            #else:
            #    sheet_us['G'+str(line)] = ('=(F' + str(line) + '/F' + str(line-1) + ') - 1')
            #    sheet_us['G'+str(line)].number_format = '00.00%'

            if sheet_us['E'+str(line)].value != 'UNDEF':
                sheet_us['F'+str(line)] = ('=AVERAGE(E' + str(line-6) + ':E' + str(line)+')')
                average_7 = sheet_us['F'+str(line)].value
                
                sheet_us['F'+str(line)].number_format = '00.00%'
            else:
                sheet_us['F'+str(line)] = 0
                sheet_us['F'+str(line)].number_format = '00.00%'

            print ('Date: ' + date + ' Infected: ' + str(infected).rjust(7)
                + ' Death: ' + str(death).rjust(5)
                #+' 7 Day Average: ' + str(average_7).rjust(5)
                )

            line += 1

    line += 1

    sheet_us['A'+str(line)] = 'Average (7-Day):'

    sheet_us['C'+str(line)] = ('=AVERAGE(C' + str(line-8) + ':C' + str(line-2)+')')
    sheet_us['C'+str(line)].number_format = '00.00%'

    sheet_us['E'+str(line)] = ('=AVERAGE(E' + str(line-8) + ':E' + str(line-2)+')')
    sheet_us['E'+str(line)].number_format = '00.00%'

    #Projections
    line += 2

    sheet_us['A'+str(line)] = ('Projections')
    line += 1

    average_7 = ((int(sheet_us['D'+str(line-5)].value)/ int(sheet_us['D'+str(line-6)].value)) + 
                (int(sheet_us['D'+str(line-6)].value)/ int(sheet_us['D'+str(line-7)].value)) + 
                (int(sheet_us['D'+str(line-7)].value)/ int(sheet_us['D'+str(line-8)].value)) + 
                (int(sheet_us['D'+str(line-8)].value)/ int(sheet_us['D'+str(line-9)].value)) + 
                (int(sheet_us['D'+str(line-9)].value)/ int(sheet_us['D'+str(line-10)].value)) + 
                (int(sheet_us['D'+str(line-10)].value)/int(sheet_us['D'+str(line-11)].value)) + 
                (int(sheet_us['D'+str(line-11)].value)/int(sheet_us['D'+str(line-12)].value))) / 7

    print()
    print()
    starting_death = int(sheet_us['D'+str(line-5)].value)
    starting_date = sheet_us['A'+str(line-5)].value

    #month, day, year = (int(x) for x in starting_date.split('-'))
    #print(type(month))
    #print(str(month) + '-' + str(day) + '-' + str(year))
    #day += 1
    #date = datetime.date(month, day, year)
    #print(date)

    print('Final Death Count (US): ' + str(starting_death))
    print('Seven Day Average Increase (US): ' + str(average_7))

    sheet_us['A'+str(line)] = ('Today')
    sheet_us['D'+str(line)] = (starting_death * average_7)
    sheet_us['E'+str(line)] = (((int(sheet_us['D'+str(line)].value)) / starting_death)-1)
    sheet_us['D'+str(line)].number_format = '0'
    sheet_us['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_us['A'+str(line)] = ('Today +1')
    sheet_us['D'+str(line)] = (sheet_us['D'+str(line-1)].value * average_7)
    sheet_us['E'+str(line)] = (((int(sheet_us['D'+str(line)].value)) / starting_death)-1)
    sheet_us['D'+str(line)].number_format = '0'
    sheet_us['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_us['A'+str(line)] = ('Today +2')
    sheet_us['D'+str(line)] = (sheet_us['D'+str(line-1)].value * average_7)
    sheet_us['E'+str(line)] = (((int(sheet_us['D'+str(line)].value)) / starting_death)-1)
    sheet_us['D'+str(line)].number_format = '0'
    sheet_us['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_us['A'+str(line)] = ('Today +3')
    sheet_us['D'+str(line)] = (sheet_us['D'+str(line-1)].value * average_7)
    sheet_us['E'+str(line)] = (((int(sheet_us['D'+str(line)].value)) / starting_death)-1)
    sheet_us['D'+str(line)].number_format = '0'
    sheet_us['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_us['A'+str(line)] = ('Today +4')
    sheet_us['D'+str(line)] = (sheet_us['D'+str(line-1)].value * average_7)
    sheet_us['E'+str(line)] = (((int(sheet_us['D'+str(line)].value)) / starting_death)-1)
    sheet_us['D'+str(line)].number_format = '0'
    sheet_us['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_us['A'+str(line)] = ('Today +5')
    sheet_us['D'+str(line)] = (sheet_us['D'+str(line-1)].value * average_7)
    sheet_us['E'+str(line)] = (((int(sheet_us['D'+str(line)].value)) / starting_death)-1)
    sheet_us['D'+str(line)].number_format = '0'
    sheet_us['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_us['A'+str(line)] = ('Today +6')
    sheet_us['D'+str(line)] = (sheet_us['D'+str(line-1)].value * average_7)
    sheet_us['E'+str(line)] = (((int(sheet_us['D'+str(line)].value)) / starting_death)-1)
    sheet_us['D'+str(line)].number_format = '0'
    sheet_us['E'+str(line)].number_format = '00.00%'

def analysis_state(sheet_name, state_name):
    print("Parsing State: "+state_name)
    death = 0
    infected = 0
    line = 1

    sheet_name['A'+str(line)] = ('Date')
    sheet_name['B'+str(line)] = ('Infected')
    sheet_name['C'+str(line)] = ('Rate')
    sheet_name['D'+str(line)] = ('Deaths')
    sheet_name['E'+str(line)] = ('Rate')
    #sheet_name['F'+str(line)] = ('Recovered')
    #sheet_name['G'+str(line)] = ('Rate')
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
                        if row[2].lower() == state_name:
                            death = death + int(row[8])
                            infected = infected + int(row[7])
                            recovered = recovered + int(row[9])
                    if row[1].lower() == 'us':
                        if row[0].lower() == state_name:
                            if row[4] != '':
                                death = death + int(row[4])
                            if row[3] != '':
                                infected = infected + int(row[3])
                            if row[5] != '':
                                recovered = recovered + int(row[5])

                    line_count += 1
                #print(f'Processed {line_count} lines.')    

            sheet_name['A'+str(line)] = date
            sheet_name['B'+str(line)] = infected
            sheet_name['D'+str(line)] = death
            #sheet_name['F'+str(line)] = recovered

            #print(type(sheet_name['B'+str(line)]))

            if sheet_name['B' + str(line-1)].value == 0:
                sheet_name['C' + str(line)] = '0'
            elif line == 2:
                sheet_name['C' + str(line)] = '0'
            else:
                sheet_name['C'+str(line)] = ('=(B' + str(line) + '/B' + str(line-1) + ') - 1')
                sheet_name['C'+str(line)].number_format = '00.00%'
            
            if sheet_name['D' + str(line-1)].value == 0:
                sheet_name['E' + str(line)] = '0'
            elif line == 2:
                sheet_name['E' + str(line)] = '0'
            else:
                sheet_name['E'+str(line)] = ('=(D' + str(line) + '/D' + str(line-1) + ') - 1')
                sheet_name['E'+str(line)].number_format = '00.00%'

            #if sheet_name['F' + str(line-1)].value == 0:
            #    sheet_name['G' + str(line)] = 'UNDEF'
            #elif line == 2:
            #    sheet_name['G' + str(line)] = 'UNDEF'
            #else:
            #    sheet_name['G'+str(line)] = ('=(F' + str(line) + '/F' + str(line-1) + ') - 1')
            #    sheet_name['G'+str(line)].number_format = '00.00%'

            if sheet_name['E'+str(line)].value != '0':
                sheet_name['F'+str(line)] = ('=AVERAGE(E' + str(line-6) + ':E' + str(line)+')')
                average_7 = sheet_name['F'+str(line)].value
                
                sheet_name['F'+str(line)].number_format = '00.00%'
            else:
                sheet_name['F'+str(line)] = 0
                sheet_name['F'+str(line)].number_format = '00.00%'

            print ('Date: ' + date + ' Infected: ' + str(infected).rjust(7)
                + ' Death: ' + str(death).rjust(5)
                #+' 7 Day Average: ' + str(average_7).rjust(5)
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

    sheet_name['A'+str(line)] = ('Today')
    sheet_name['D'+str(line)] = (starting_death * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_name['A'+str(line)] = ('Today +1')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_name['A'+str(line)] = ('Today +2')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_name['A'+str(line)] = ('Today +3')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_name['A'+str(line)] = ('Today +4')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_name['A'+str(line)] = ('Today +5')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_name['A'+str(line)] = ('Today +6')
    sheet_name['D'+str(line)] = (sheet_name['D'+str(line-1)].value * average_7)
    sheet_name['E'+str(line)] = (((int(sheet_name['D'+str(line)].value)) / starting_death)-1)
    sheet_name['D'+str(line)].number_format = '0'
    sheet_name['E'+str(line)].number_format = '00.00%'

    print()
    print()

analysis_us()
analysis_state(sheet_ga, 'georgia')
analysis_state(sheet_sc, 'south carolina')
analysis_state(sheet_nc, 'north carolina')
analysis_state(sheet_ny, 'new york')
analysis_state(sheet_wa, 'washington')
analysis_state(sheet_ca, 'california')


date_object = datetime.date.today()

workbook.save(filename=cwd_dir + '\\' + str(date_object) + wb_name)

print()
print("Complete")
print()
print()
#input("Press Enter to continue...")

