from lxml import html
import xlsxwriter

# code to get crn, dept, coursenumber, section, coursecredits, and coursetitle into a excel doc

workbook = xlsxwriter.Workbook('CourseCatalogwTeachers.xlsx')
worksheet = workbook.add_worksheet()

with open('coursesSpring2018.txt') as myFile:
    data = myFile.read()

classes = html.fromstring(data)

# note that classes that have two different times will have a \xa0 after them
crn = classes.xpath('//table[@class="datadisplaytable"]//tr//td[2]//text()')
dept = classes.xpath('//table[@class="datadisplaytable"]//tr//td[3]//text()')
courseNumber = classes.xpath('//table[@class="datadisplaytable"]//tr//td[4]//text()')
section = classes.xpath('//table[@class="datadisplaytable"]//tr//td[5]//text()')
courseCredits = classes.xpath('//table[@class="datadisplaytable"]//tr//td[7]//text()')
courseTitle = classes.xpath('//table[@class="datadisplaytable"]//tr//td[8]//text()')
days = classes.xpath('//table[@class="datadisplaytable"]//tr//td[9]//text()') # Need to sort out classes that have different times
time = classes.xpath('//table[@class="datadisplaytable"]//tr//td[10]//text()') # Need way to sort out classes that have two different times
prof = classes.xpath('//table[@class="datadisplaytable"]//tr//td[20]//text()') # get rid of the (', 'P', ')

profNew = []
col = 0
index = 0
row = 0


for item in crn:
    if crn[index] == " ":
        index += 1
        continue
    else:
        worksheet.write(row, col, crn[index])
        row += 1
        index += 1

row = 0
col += 1
index = 0

for item in dept:
    if crn[index] == " ":
        index += 1
        continue
    else:
        worksheet.write(row, col, dept[index])
        row += 1
        index += 1

row = 0
col += 1
index = 0

for item in courseNumber:
    if crn[index] == " ":
        index += 1
        continue
    else:
        worksheet.write(row, col, courseNumber[index])
        row += 1
        index += 1

row = 0
col += 1
index = 0

for item in section:
    if crn[index] == " ":
        index += 1
        continue
    else:
        worksheet.write(row, col, section[index])
        row += 1
        index += 1

row = 0
col += 1
index = 0

for item in courseCredits:
    if crn[index] == " ":
        index += 1
        continue
    else:
        worksheet.write(row, col, courseCredits[index])
        row += 1
        index += 1

row = 0
col += 1
index = 0

for item in courseTitle:
    if crn[index] == " ":
        index += 1
        continue
    else:
        worksheet.write(row, col, courseTitle[index])
        row += 1
        index += 1

row = 0
col += 1
index = 0
reference = 0
prof[:] = [x for x in prof if x != 'P']
prof[:] = [x for x in prof if x != ')']
for item in prof:
    #Need to add 3rd reference variable to make the list compatible with crn[reference]
    if crn[reference] == ' ':
        reference += 1
        index += 1
        continue

    elif '),' in prof[index]:
        index += 1
        continue

    else:
        profNew.append(prof[index].replace("(", ""))
        print(prof[index])
        worksheet.write(row, col, profNew[row])
        row += 1
        index += 1
        reference += 1

row = 0
col += 1
index = 0
shift = 0

for item in days:
    if crn[index] == " ":
        row -= 1
        shift += 2
        worksheet.write(row, col + shift, days[index])
        row += 1
        index += 1
        continue
    else:
        worksheet.write(row,col, days[index])
        row += 1
        index += 1
    shift = 0

row = 0
col += 1
index = 0

for item in time:
    if crn[index] == " ":
        row -= 1
        shift += 2
        worksheet.write(row, col + shift, time[index])
        row += 1
        index += 1
        continue

    else:
        worksheet.write(row, col, time[index])
        row += 1
        index += 1

    shift = 0
row = 0
col += 1
index = 0


col += 1
index = 0

workbook.close()