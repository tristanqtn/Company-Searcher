# ================================================ IMPORTS
# import xlsxwriter module
import xlsxwriter

# import Google Search API
from googlesearch import search

# import date and time service
from datetime import datetime

# ================================================ BEGINNING PROGRAM
print("\n ===============================================================================================")
print("    _____                                            _____                     _")
print("   / ____|                                          / ____|                   | |")
print("  | |     ___  _ __ ___  _ __   __ _ _ __  _   _   | (___   ___  __ _ _ __ ___| |__   ___ _ __")
print("  | |    / _ \| '_ ` _ \| '_ \ / _` | '_ \| | | |   \___ \ / _ \/ _` | '__/ __| '_ \ / _ \ '__|")
print("  | |___| (_) | | | | | | |_) | (_| | | | | |_| |   ____) |  __/ (_| | | | (__| | | |  __/ |")
print("   \_____\___/|_| |_| |_| .__/ \__,_|_| |_|\__, |  |_____/ \___|\__,_|_|  \___|_| |_|\___|_|")
print("                        | |                 __/ |")
print("                        |_|                |___/  ")
print(" ===============================================================================================")

# counting number of line
with open("./data.txt", 'r') as fp:
    for count, line in enumerate(fp):
        pass
# closing files
fp.close()
# increments final count
count += 1
print('\nNumber of companies written in file: ', count)
# asking for number of research per company
search_number = input("Number of search results wanted: ")
# casting into int
search_number = int(search_number)

# open file containing company names
f = open("./data.txt", "r")


# get the current date


# datetime object containing current date and time
now = datetime.now()
dt_string = now.strftime('%d-%m-%Y_%H-%M-%S')

# generating name
output_file_name = dt_string
output_file_name += '_cs_output'
output_file_name += '.xlsx'
print("Output file name: " + output_file_name)
# creating workbook
workbook = xlsxwriter.Workbook(output_file_name)
# creation worksheet
worksheet = workbook.add_worksheet()

row = 0

# for each company name
for x in range(count):
    # read company name
    query = f.readline()
    print("\n\n     >> Searching for: " + query)
    column = 0
    # writing the name in the worksheet
    worksheet.write(row, column, query)
    # SEARCHING HERE
    for j in search(query, tld="co.in", num=search_number, stop=search_number, pause=2):
        column += 1
        # writing searching result in worksheet
        worksheet.write(row, column, j)
        print(j)
    row += 1
# closing files
f.close()
workbook.close()
print("\n===============================================================================================")

# ================================================ PROGRAM END
