import json
import re
from openpyxl import load_workbook

src_filename = "AquíNecesitamos.xlsx"
dst_filename = "output.json"
src_sheet = ["URGENCIAS Y SOLICITUDES POR ZON"]
sheet_index = 0;
MAX_TWEET_CHARACTERS = 140
DATA_MIN_ROW = 6
DATA_MAX_COL = 9
INFO_HASH_TAG = "#infoverificada19S"

# Data label
URGENT_LEVEL = {"alto":"URGE", "alta":"URGE", "medio":"SeNecesita", "media":"SeNecesita", "bajo":"SeNecesita", "baja":"SeNecesita"}

def main():
    tweetList = [];
    wb = load_workbook(src_filename)
    ws = wb[src_sheet[sheet_index]]
    for row in ws.iter_rows(min_row=DATA_MIN_ROW, max_col=DATA_MAX_COL):
        row_data = []
        for cell in row:
            row_data.append(cell.value);
        tweet = generateText(row_data)
        tweetList.append(convertTweetToJson(tweet))
    jsonData = createJson(tweetList)
    print("Generated: " + str(len(tweetList)) + " tweets")
    saveFile(jsonData, dst_filename)
    print("---Finish saving generated text to: " + dst_filename + "---")

#UrgentLevel (X)
#NEED BRIGADISTS (X)
#MOST IMPORTANT REQUIRED
#ADMITTED
#NOT REQUIRED
#ADDRESS
#ZONE
#DETAIL/SOURCE
#UPGRADE
def generateText(row):
    tweet = ""

    # Time
    if (row[8] is not None):
        time = (row[8])[5:] # remove year
        tweet += time

    # Urgent level
    if (row[0] is not None):
        urgent = row[0].lower()
        urgent = urgent.rstrip()
        urgent = urgent.lstrip()
        tweet += " " + URGENT_LEVEL[urgent]+""

####    # Need Brigadists
####    needHelp = (row[1].lower() == "si") or (row[1].lower() == "sí")
####    if (needHelp):
####        tweet += ", NECESITAN BRIGADISTAS"
##
##
    # Hash Tag
    tweet += INFO_HASH_TAG

    # Address
    if (row[5] is not None):
        address = row[5]
        address = getAddress(address)
        tweet += " Rescate en " + address

    # Zone
    if (row[6] is not None):
        zone = row[6]
        tweet += " " + zone

    return tweet

#=HYPERLINK("https://goo.gl/maps/RLuTSzXwLWm","Eje Central 806, Esquina Niños Héroes")'
def getAddress(address):
    address = str(address)
    formatText = '=HYPERLINK'
    hasLink = (address[0:len(formatText)]) == formatText
    if (hasLink):
        address = address[len(formatText):]
        return "TEST ADDRESS"
    else:
        return address

def checkTweetLength(text):
    return (len(text) < MAX_TWEET_CHARACTERS)

def convertTweetToJson(tweet):
    data = {"text": tweet}
    return json.dumps(data)

def createJson(tweetList):
    output = "["
    for i in range(len(tweetList)):
        if (i == len(tweetList)-1):
            output += tweetList[i]
        else:
            output += tweetList[i] + ", "
    output += "]"
    return output

def saveFile(data, filename):
    file = open(filename, 'w')
    file.write(data)
    file.close

main()
