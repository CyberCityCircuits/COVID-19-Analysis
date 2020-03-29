#Data Source: https://github.com/CSSEGISandData/COVID-19
#COVID-19 Aggregation Script by Cyber City Circuits

import csv, os
import openpyxl
from openpyxl import Workbook
import datetime

wb_name = '_aggregated.xlsx'
data_dir = 'C:\\Users\\DREAM\\Documents\\GitHub\\COVID-19 Scripts\\data\\'

death = 0
infected = 0
line = 1

workbook = Workbook()
sheet = workbook.active
sheet['A'+str(line)] = ('Date')
sheet['B'+str(line)] = ('Infected')
sheet['C'+str(line)] = ('Rate')
sheet['D'+str(line)] = ('Deaths')
sheet['E'+str(line)] = ('Rate')
#sheet['F'+str(line)] = ('Recovered')
#sheet['G'+str(line)] = ('Rate')
sheet['F'+str(line)] = ('Avaerage (7-Day)')

sheet.column_dimensions['A'].width = 10
sheet.column_dimensions['F'].width = 16

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

        sheet['A'+str(line)] = date
        sheet['B'+str(line)] = infected
        sheet['D'+str(line)] = death
        #sheet['F'+str(line)] = recovered

        #print(type(sheet['B'+str(line)]))

        if sheet['B' + str(line-1)].value == 0:
            sheet['C' + str(line)] = 'UNDEF'
        elif line == 2:
            sheet['C' + str(line)] = 'UNDEF'
        else:
            sheet['C'+str(line)] = ('=(B' + str(line) + '/B' + str(line-1) + ') - 1')
            sheet['C'+str(line)].number_format = '00.00%'
        
        if sheet['D' + str(line-1)].value == 0:
            sheet['E' + str(line)] = 'UNDEF'
        elif line == 2:
            sheet['E' + str(line)] = 'UNDEF'
        else:
            sheet['E'+str(line)] = ('=(D' + str(line) + '/D' + str(line-1) + ') - 1')
            sheet['E'+str(line)].number_format = '00.00%'

        #if sheet['F' + str(line-1)].value == 0:
        #    sheet['G' + str(line)] = 'UNDEF'
        #elif line == 2:
        #    sheet['G' + str(line)] = 'UNDEF'
        #else:
        #    sheet['G'+str(line)] = ('=(F' + str(line) + '/F' + str(line-1) + ') - 1')
        #    sheet['G'+str(line)].number_format = '00.00%'

        if sheet['E'+str(line)].value != 'UNDEF':
            sheet['F'+str(line)] = ('=AVERAGE(E' + str(line-6) + ':E' + str(line)+')')
            average_7 = sheet['F'+str(line)].value
            
            sheet['F'+str(line)].number_format = '00.00%'
        else:
            sheet['F'+str(line)] = 0
            sheet['F'+str(line)].number_format = '00.00%'

        print ('Date: ' + date + ' Infected: ' + str(infected).rjust(7)
              + ' Death: ' + str(death).rjust(5)
              #+' 7 Day Average: ' + str(average_7).rjust(5)
              )

        line += 1

line += 1

sheet['A'+str(line)] = 'Average (7-Day):'

sheet['C'+str(line)] = ('=AVERAGE(C' + str(line-8) + ':C' + str(line-2)+')')
sheet['C'+str(line)].number_format = '00.00%'

sheet['E'+str(line)] = ('=AVERAGE(E' + str(line-8) + ':E' + str(line-2)+')')
sheet['E'+str(line)].number_format = '00.00%'

#Projections
line += 2

sheet['A'+str(line)] = ('Projections')
line += 1

average_7 = ((int(sheet['D'+str(line-5)].value)/ int(sheet['D'+str(line-6)].value)) + 
             (int(sheet['D'+str(line-6)].value)/ int(sheet['D'+str(line-7)].value)) + 
             (int(sheet['D'+str(line-7)].value)/ int(sheet['D'+str(line-8)].value)) + 
             (int(sheet['D'+str(line-8)].value)/ int(sheet['D'+str(line-9)].value)) + 
             (int(sheet['D'+str(line-9)].value)/ int(sheet['D'+str(line-10)].value)) + 
             (int(sheet['D'+str(line-10)].value)/int(sheet['D'+str(line-11)].value)) + 
             (int(sheet['D'+str(line-11)].value)/int(sheet['D'+str(line-12)].value))) / 7



print()
print()
starting_death = int(sheet['D'+str(line-5)].value)
starting_date = sheet['A'+str(line-5)].value

#month, day, year = (int(x) for x in starting_date.split('-'))
#print(type(month))
#print(str(month) + '-' + str(day) + '-' + str(year))
#day += 1
#date = datetime.date(month, day, year)
#print(date)

print('Final Death Count: ' + str(starting_death))
print('Seven Day Average Increase: ' + str(average_7))

sheet['A'+str(line)] = ('Today')
sheet['D'+str(line)] = (starting_death * average_7)
sheet['E'+str(line)] = (((int(sheet['D'+str(line)].value)) / starting_death)-1)
sheet['D'+str(line)].number_format = '0'
sheet['E'+str(line)].number_format = '00.00%'
line += 1

sheet['A'+str(line)] = ('Day +2')
sheet['D'+str(line)] = (sheet['D'+str(line-1)].value * average_7)
sheet['E'+str(line)] = (((int(sheet['D'+str(line)].value)) / starting_death)-1)
sheet['D'+str(line)].number_format = '0'
sheet['E'+str(line)].number_format = '00.00%'
line += 1

sheet['A'+str(line)] = ('Day +3')
sheet['D'+str(line)] = (sheet['D'+str(line-1)].value * average_7)
sheet['E'+str(line)] = (((int(sheet['D'+str(line)].value)) / starting_death)-1)
sheet['D'+str(line)].number_format = '0'
sheet['E'+str(line)].number_format = '00.00%'
line += 1

sheet['A'+str(line)] = ('Day +4')
sheet['D'+str(line)] = (sheet['D'+str(line-1)].value * average_7)
sheet['E'+str(line)] = (((int(sheet['D'+str(line)].value)) / starting_death)-1)
sheet['D'+str(line)].number_format = '0'
sheet['E'+str(line)].number_format = '00.00%'
line += 1

sheet['A'+str(line)] = ('Day +5')
sheet['D'+str(line)] = (sheet['D'+str(line-1)].value * average_7)
sheet['E'+str(line)] = (((int(sheet['D'+str(line)].value)) / starting_death)-1)
sheet['D'+str(line)].number_format = '0'
sheet['E'+str(line)].number_format = '00.00%'
line += 1

sheet['A'+str(line)] = ('Day +6')
sheet['D'+str(line)] = (sheet['D'+str(line-1)].value * average_7)
sheet['E'+str(line)] = (((int(sheet['D'+str(line)].value)) / starting_death)-1)
sheet['D'+str(line)].number_format = '0'
sheet['E'+str(line)].number_format = '00.00%'
line += 1

sheet['A'+str(line)] = ('Day +7')
sheet['D'+str(line)] = (sheet['D'+str(line-1)].value * average_7)
sheet['E'+str(line)] = (((int(sheet['D'+str(line)].value)) / starting_death)-1)
sheet['D'+str(line)].number_format = '0'
sheet['E'+str(line)].number_format = '00.00%'

workbook.save(filename=data_dir + wb_name)

print("Complete")
