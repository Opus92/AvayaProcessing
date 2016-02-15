#
# Process and append data from many CSVs to
# a handful of composite CSVs.
#

import os
import shutil

# Set source and working directory
src = r'U:\Projects\Telecom\RawData'
os.chdir(src)

# Set destination directory
dst0 = r'U:\Projects\Telecom\ProcessedData'
dst1 = r'P:\IT-SS-OA'

# Create and open composite files for processed data
a = open(r'U:\Projects\Telecom\_avReports.csv', 'a')
i = open(r'U:\Projects\Telecom\_ithdReports.csv', 'a')
l = open(r'U:\Projects\Telecom\_lmsReports.csv', 'a')
o = open(r'U:\Projects\Telecom\_opReports.csv', 'a')
w = open(r'U:\Projects\Telecom\_wsReports.csv', 'a')
v = open(r'U:\Projects\Telecom\_vdnReports.csv', 'a')
log = open(r'U:\Projects\Telecom\_loginReports.csv', 'a')

# Basic formatting and text cleanup for Split reports
def splitClean(text):
    text = text.replace(',\"-\",',',')
    text = text.replace('\"','')
    text = text.replace('0AM','0 AM')
    newText = text.replace('0PM','0 PM')

    return newText

# Basic formatting and text cleanup for VDN reports
def vdnClean(text):
    text = text.replace('\"','')
    text = text.replace(',-','')
    text = text.replace('M,475,','M,')
    text = text.replace('0AM','0 AM')
    newText = text.replace('0PM','0 PM')
  
    return newText

# Basic formatting and text cleanup for Login reports
def logClean(text):
    text = text.replace('\" \",','')
    text = text.replace('\"\",','')
    text = text.replace('\"','')
    text = text.replace('AM',' AM')
    newText = text.replace('PM',' PM')

    return newText

# Inserts the report date in each line of the Login report
def logDate(text, date):
    text = text.split(',')
    text.insert(3, date)

    newText = ','.join(text)
    return newText

# Change date to yyyy/mm/dd format
def dateClean(text):
    text = text.split('/')

    if len(text[1]) < 2:
        text[1] = '0' + text[1]

    if len(text[0]) < 2:
        text[0] = '0' + text[0]
    
    newText = text[2] + '/' + text[0] + '/' + text[1]

    return newText

# Formats text of recent CSV files
def timeClean(text):
    text = text.split(',')

    # Append AM/PM to start times
    ampm = text[1]
    n = 0

    if text[1] == '12:00 AM':
        text[0] = text[0] + ' PM'
    elif text[1] == '12:00 PM':
        text[0] = text[0] + ' AM'
    else:
        text[0] = text[0] + ' ' + ampm[-2:]

    # Convert timespans to hh:mm:ss format
    for item in text:
        strItem = str(item)
        
        if strItem.find(':') > -1 and strItem.find('M') < 0:
            if len(strItem) is 3 or strItem[3:] == '\n':
                newItem = '00:00' + strItem
                text[n] = newItem

            elif len(strItem) is 4 or strItem[4:] == '\n':
                newItem = '00:0' + strItem
                text[n] = newItem

            elif len(strItem) is 5 or strItem[5:] == '\n':
                splitTime = strItem.split(':')

                # Convert time durations of more than one hour to hh:mm:ss format
                if int(splitTime[0]) >= 60:
                    intHour = int(int(splitTime[0]) / 60)
                    intMinute = int(splitTime[0]) % 60

                    if intMinute < 10:
                        strMinute = '0' + str(intMinute)
                        strHour = str(intHour)

                    else:
                        strMinute = str(intMinute)
                        strHour = str(intHour)

                    splitTime[0] = strHour
                    splitTime.insert(1, strMinute)
                    strItem = (':').join(splitTime)
                
                    newItem = strItem

                else:
                    newItem = '00:' + strItem
                    
                text[n] = newItem

        else:
            pass

        n = n + 1

    newText = (',').join(text)
    return newText


# Writes column headers. This only needs to be done per composite file.
##splitHeader = 'Date,StartTime,StopTime,Avg Speed Ans,Avg Aban Time,ACD Calls,Avg ACD Time,Avg ACW Time,Aban Calls,Max Delay,Flow In,Flow Out,Extn Out Calls,Avg Extn Out Time,Dequeued Calls,Avg Time to Dequeue,Percent ACD Time,Percent Ans Calls,Avg Pos Staff,Calls Per Pos\n'
##vdnHeader = 'Date,StartTime,StopTime,Vector,Inbound Calls,Flow In,ACD Calls,Avg Speed Ans,Avg ACD Time,Avg ACW Time,Main ACD Calls,Backup ACD Calls,Connect Calls,Avg Connect Time,Aban Calls,Avg Aban Time,Forced Busy Calls,Forced Disc Calls,Flow Out,Avg VDN Time\n'
##logHeader = 'Dept,Agent,Extn,LoginTime,LoginDate,LogoutTime,LogoutDate\n'
##
##a.write(splitHeader)
##i.write(splitHeader)
##l.write(splitHeader)
##o.write(splitHeader)
##w.write(splitHeader)
##v.write(vdnHeader)
##log.write(logHeader)

# Loop through files in a directory and process them according to the first letter of the filename
for dirName, subdirList, fileList in os.walk(src):
    for filename in fileList:     
        with open (os.path.join(os.getcwd(), dirName, filename), 'r') as f:
            charTest = filename[0:1]
            
            for line in f:
                if charTest[0] in ['_', 'S']:
                    pass
            
                elif charTest in ('AV','IT','LM','Op','Wo'):
                    line = splitClean(line)
                    lineTest = line[0]

                    if lineTest in ['H', 'T', 'S', 'd', '\n']:
                        pass
                    elif lineTest in ['D']:
                        theDate = line[6:len(line) - 1]
                        theDate = dateClean(theDate)
                    else:
                        line = splitClean(line)
                        line = timeClean(line)
                        line = theDate + ',' + line

                        if charTest == 'AV':
                            a.write(line)

                        elif charTest == 'IT':
                            i.write(line)

                        elif charTest == 'LM':
                            l.write(line)

                        elif charTest == 'Op':
                            o.write(line)

                        elif charTest == 'Wo':
                            w.write(line)
                        
                        else:
                            print("Something went wrong in writing lines to files for", filename)
                            input('Press any key to continue.')
                    
                elif charTest in ('VD'):
                    line = vdnClean(line)
                    lineTest = line[0]

                    if lineTest in ['H', 'T', 'V', 'd', '\n']:
                        pass
                    
                    elif lineTest in ['D']:
                        theDate = line[6:len(line) - 1]
                        theDate = dateClean(theDate)
                        
                    else:
                        line = vdnClean(line)
                        line = timeClean(line)
                        line = theDate + ',' + line

                        v.write(line)
                        
                elif charTest == 'Te':
                    pass

                elif charTest in ['20']:
                    line = logClean(line)
                    lineTest = line[0]

                    if lineTest in ['A', 'H', 'T', 'd', '\n']:
                        pass

                    elif lineTest in ['S']:
                        theDept = line[7:len(line) - 1]

                    elif lineTest in ['D']:
                        theDate = line[6:len(line) - 1]

                    else:
                        line = logDate(line, theDate)
                        line = theDept + ',' + line                     
                        log.write(line)
                
                else:
                    print('Something went wrong in choosing a processing path for', filename)
                    input('Press any key to continue.')

        # Close unprocessed file, move it to another directory, and notify the user
        f.close()
        shutil.move((os.path.join(os.getcwd(), dirName, filename)), dst0)
        print ('Processing for', filename, 'is complete.')

input('Press any key to continue.')

# Close composite files
a.close()
i.close()
l.close()
o.close()
w.close()
v.close()
log.close()

# Copy composite files to other directories
##os.chdir(dst1)
##for dirName, subdirList, fileList in os.walk(dst1):
##    for filename in fileList:
##        charTest = filename[0]
##
##        if charTest is '_':
##            shutil.copy((os.path.join(os.getcwd(), dirName, filename)), dst1)
##        else:
##            pass
