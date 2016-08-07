import os, csv
import re
report = {}
regex = re.compile(("([a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`"
                        "{|}~-]+)*(@|\sat\s)(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(\.|"
                        "\sdot\s))+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)"))
def getEmailIds():
    print "Getting list of Students.."
    for path,dirs,files in os.walk('.'):
        for date in dirs:
            for path,dirs1,files in os.walk(date):
                for filename in files:
                    email = re.findall(regex, filename)
                    for i in email:
                        report[i[0]] = {}
                        break
    print "Total Students: "+str(len(report))

def createDataObject(date):
    arr = []
    for path, dirs, files in os.walk(date):
        print "\nFolder: "+date+"                Size: "+str(len(files))+""
        print "-----------------------------------------------"
        print "Scanning "+str(len(files))+" Files...."
        for filename in files:
            if 'log' in filename:
                email = re.findall(regex, filename)
                for i in email:
                    arr.append(i[0])
                    break
                with open(date+filename) as f:
                    keystrokes = 0
                    mousemoves = -1
                    total = 0
                    content = f.readlines()
                    for i in content:
                        total += 1
                        if 'mouse' not in str(i):
                            keystrokes += 1
                        mousemoves = total - keystrokes
                    for i in email:
                        report[i[0]][date] = [keystrokes, mousemoves]
    for entry in list(set(report.keys())-set(arr)):
        report[entry][date] = ["None", "None"]
    average = 0
    for email in report:
        if report[email][date][0] != "None":
            average += report[email][date][0]
        else:
            average += 0
    average = average/len(report)
    return average
    

def generateExcel(dataobj):
    print "\n\nImporting the data into Excel Workbook.."
    import xlsxwriter
    workbook = xlsxwriter.Workbook('WatchdogDataCatalog.xlsx')
    worksheet = workbook.add_worksheet('Main Sheet')
    worksheet2 = workbook.add_worksheet('Key Strokes Report')
    worksheet.set_column('A:A', 30)
    worksheet2.set_column('A:A', 30)
    row = 2
    for email in dataobj:
        totavg = getOverallAverage(dataobj[email])
        worksheet.write(row, 0, email)
        worksheet2.write(row, 0, email)
        worksheet2.write(row, 1, totavg)
        col=2
        col1 = 2
        for data in dataobj[email]:
            worksheet.write(0, col, data[:-1])
            worksheet2.write(1,1,"Total")
            worksheet2.write(0, col1, data[:-1])
            worksheet.write(1,col,"Key Strokes")
            worksheet2.write(1,col1,"Key Strokes Average")
            worksheet.write(1,col+1,"Mouse Events")
            worksheet.write(row,col,dataobj[email][data][0])
            worksheet2.write(row,col1,dataobj[email][data][2])
            worksheet.write(row,col+1,dataobj[email][data][1])
            worksheet.merge_range(0,col,0,col+1,data[:-1])
            col+=2
            col1+=1
        row+=1
        col+=1

    print "Setting things up... Done."
    workbook.close()

def getKeystrokeCodes():
    for path,dirs,files in os.walk('.'):
        print "Total Folders: "+str(len(dirs))+""
        for i in dirs:
            averagekeystroke = createDataObject(i+"/")
            print "\nDay: "+i
            print "Average number of Key Strokes: "+str(averagekeystroke)
            aa = 0
            ba = 0
            na = 0
            for email in report:
                if report[email][i+"/"][0] == "None":
                    report[email][i+"/"].append(0)
                    na+=1
                elif report[email][i+"/"][0] > averagekeystroke:
                    report[email][i+"/"].append(2)
                    aa+=1
                else:
                    report[email][i+"/"].append(1)
                    ba+=1
            print "Students above Average: "+str(aa)
            print "Students below Average: "+str(ba)
            print "Inactive Students: "+str(na)
        break

def getOverallAverage(email):
        totalaverage = 0
        for date in email:
            totalaverage += email[date][2]
        return totalaverage

getEmailIds()
getKeystrokeCodes()
generateExcel(report)