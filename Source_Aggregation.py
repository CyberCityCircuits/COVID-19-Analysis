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

    date_object = datetime.date.today()

    workbook.save(filename=cwd_dir + '\\' + str(date_object) + wb_name)


def analysis_ga():
    death = 0
    infected = 0
    line = 1

    sheet_ga['A'+str(line)] = ('Date')
    sheet_ga['B'+str(line)] = ('Infected')
    sheet_ga['C'+str(line)] = ('Rate')
    sheet_ga['D'+str(line)] = ('Deaths')
    sheet_ga['E'+str(line)] = ('Rate')
    #sheet_ga['F'+str(line)] = ('Recovered')
    #sheet_ga['G'+str(line)] = ('Rate')
    sheet_ga['F'+str(line)] = ('Avaerage (7-Day)')

    sheet_ga.column_dimensions['A'].width = 10
    sheet_ga.column_dimensions['F'].width = 16

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
                        if row[2].lower() == 'georgia':
                            death = death + int(row[8])
                            infected = infected + int(row[7])
                            recovered = recovered + int(row[9])
                    if row[1].lower() == 'us':
                        if row[0].lower() == 'georgia':
                            if row[4] != '':
                                death = death + int(row[4])
                            if row[3] != '':
                                infected = infected + int(row[3])
                            if row[5] != '':
                                recovered = recovered + int(row[5])

                    line_count += 1
                #print(f'Processed {line_count} lines.')    

            sheet_ga['A'+str(line)] = date
            sheet_ga['B'+str(line)] = infected
            sheet_ga['D'+str(line)] = death
            #sheet_ga['F'+str(line)] = recovered

            #print(type(sheet_ga['B'+str(line)]))

            if sheet_ga['B' + str(line-1)].value == 0:
                sheet_ga['C' + str(line)] = 'UNDEF'
            elif line == 2:
                sheet_ga['C' + str(line)] = 'UNDEF'
            else:
                sheet_ga['C'+str(line)] = ('=(B' + str(line) + '/B' + str(line-1) + ') - 1')
                sheet_ga['C'+str(line)].number_format = '00.00%'
            
            if sheet_ga['D' + str(line-1)].value == 0:
                sheet_ga['E' + str(line)] = 'UNDEF'
            elif line == 2:
                sheet_ga['E' + str(line)] = 'UNDEF'
            else:
                sheet_ga['E'+str(line)] = ('=(D' + str(line) + '/D' + str(line-1) + ') - 1')
                sheet_ga['E'+str(line)].number_format = '00.00%'

            #if sheet_ga['F' + str(line-1)].value == 0:
            #    sheet_ga['G' + str(line)] = 'UNDEF'
            #elif line == 2:
            #    sheet_ga['G' + str(line)] = 'UNDEF'
            #else:
            #    sheet_ga['G'+str(line)] = ('=(F' + str(line) + '/F' + str(line-1) + ') - 1')
            #    sheet_ga['G'+str(line)].number_format = '00.00%'

            if sheet_ga['E'+str(line)].value != 'UNDEF':
                sheet_ga['F'+str(line)] = ('=AVERAGE(E' + str(line-6) + ':E' + str(line)+')')
                average_7 = sheet_ga['F'+str(line)].value
                
                sheet_ga['F'+str(line)].number_format = '00.00%'
            else:
                sheet_ga['F'+str(line)] = 0
                sheet_ga['F'+str(line)].number_format = '00.00%'

            print ('Date: ' + date + ' Infected: ' + str(infected).rjust(7)
                + ' Death: ' + str(death).rjust(5)
                #+' 7 Day Average: ' + str(average_7).rjust(5)
                )

            line += 1

    line += 1

    sheet_ga['A'+str(line)] = 'Average (7-Day):'

    sheet_ga['C'+str(line)] = ('=AVERAGE(C' + str(line-8) + ':C' + str(line-2)+')')
    sheet_ga['C'+str(line)].number_format = '00.00%'

    sheet_ga['E'+str(line)] = ('=AVERAGE(E' + str(line-8) + ':E' + str(line-2)+')')
    sheet_ga['E'+str(line)].number_format = '00.00%'

    #Projections
    line += 2

    sheet_ga['A'+str(line)] = ('Projections')
    line += 1

    average_7 = ((int(sheet_ga['D'+str(line-5)].value)/ int(sheet_ga['D'+str(line-6)].value)) + 
                (int(sheet_ga['D'+str(line-6)].value)/ int(sheet_ga['D'+str(line-7)].value)) + 
                (int(sheet_ga['D'+str(line-7)].value)/ int(sheet_ga['D'+str(line-8)].value)) + 
                (int(sheet_ga['D'+str(line-8)].value)/ int(sheet_ga['D'+str(line-9)].value)) + 
                (int(sheet_ga['D'+str(line-9)].value)/ int(sheet_ga['D'+str(line-10)].value)) + 
                (int(sheet_ga['D'+str(line-10)].value)/int(sheet_ga['D'+str(line-11)].value)) + 
                (int(sheet_ga['D'+str(line-11)].value)/int(sheet_ga['D'+str(line-12)].value))) / 7

    print()
    print()
    starting_death = int(sheet_ga['D'+str(line-5)].value)
    starting_date = sheet_ga['A'+str(line-5)].value

    #month, day, year = (int(x) for x in starting_date.split('-'))
    #print(type(month))
    #print(str(month) + '-' + str(day) + '-' + str(year))
    #day += 1
    #date = datetime.date(month, day, year)
    #print(date)

    print('Final Death Count (US): ' + str(starting_death))
    print('Seven Day Average Increase (US): ' + str(average_7))

    sheet_ga['A'+str(line)] = ('Today')
    sheet_ga['D'+str(line)] = (starting_death * average_7)
    sheet_ga['E'+str(line)] = (((int(sheet_ga['D'+str(line)].value)) / starting_death)-1)
    sheet_ga['D'+str(line)].number_format = '0'
    sheet_ga['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_ga['A'+str(line)] = ('Today +1')
    sheet_ga['D'+str(line)] = (sheet_ga['D'+str(line-1)].value * average_7)
    sheet_ga['E'+str(line)] = (((int(sheet_ga['D'+str(line)].value)) / starting_death)-1)
    sheet_ga['D'+str(line)].number_format = '0'
    sheet_ga['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_ga['A'+str(line)] = ('Today +2')
    sheet_ga['D'+str(line)] = (sheet_ga['D'+str(line-1)].value * average_7)
    sheet_ga['E'+str(line)] = (((int(sheet_ga['D'+str(line)].value)) / starting_death)-1)
    sheet_ga['D'+str(line)].number_format = '0'
    sheet_ga['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_ga['A'+str(line)] = ('Today +3')
    sheet_ga['D'+str(line)] = (sheet_ga['D'+str(line-1)].value * average_7)
    sheet_ga['E'+str(line)] = (((int(sheet_ga['D'+str(line)].value)) / starting_death)-1)
    sheet_ga['D'+str(line)].number_format = '0'
    sheet_ga['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_ga['A'+str(line)] = ('Today +4')
    sheet_ga['D'+str(line)] = (sheet_ga['D'+str(line-1)].value * average_7)
    sheet_ga['E'+str(line)] = (((int(sheet_ga['D'+str(line)].value)) / starting_death)-1)
    sheet_ga['D'+str(line)].number_format = '0'
    sheet_ga['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_ga['A'+str(line)] = ('Today +5')
    sheet_ga['D'+str(line)] = (sheet_ga['D'+str(line-1)].value * average_7)
    sheet_ga['E'+str(line)] = (((int(sheet_ga['D'+str(line)].value)) / starting_death)-1)
    sheet_ga['D'+str(line)].number_format = '0'
    sheet_ga['E'+str(line)].number_format = '00.00%'
    line += 1

    sheet_ga['A'+str(line)] = ('Today +6')
    sheet_ga['D'+str(line)] = (sheet_ga['D'+str(line-1)].value * average_7)
    sheet_ga['E'+str(line)] = (((int(sheet_ga['D'+str(line)].value)) / starting_death)-1)
    sheet_ga['D'+str(line)].number_format = '0'
    sheet_ga['E'+str(line)].number_format = '00.00%'

    date_object = datetime.date.today()

    workbook.save(filename=cwd_dir + '\\' + str(date_object) + wb_name)

analysis_us()
analysis_ga()

print()
print("Complete")
print()
print()
#input("Press Enter to continue...")

