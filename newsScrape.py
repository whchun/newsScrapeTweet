import json
import re
from openpyxl import load_workbook

src_filename = "AquíNecesitamos.xlsx"
src_sheet = ["URGENCIAS Y SOLICITUDES POR ZON"]
sheet_index = 0;
MAX_TWEET_CHARACTERS = 140
DATA_MIN_ROW = 6
DATA_MAX_COL = 9

# Data label
URGENT_LEVEL = {"alto":"URGE", "medio":"SeNecesita", "bajo":"SeNecesita"}

def main():
    wb = load_workbook(src_filename)
    ws = wb[src_sheet[sheet_index]]
    for row in ws.iter_rows(min_row=DATA_MIN_ROW, max_col=DATA_MAX_COL):
        row_data = []
        for cell in row:
            row_data.append(cell.value);
        generateText(row_data)

def generateText(row):
    tweet = ""

    # Time
    time = (row[8])[5:] # remove year
    tweet += time
    
    # Urgent level
    urgent = row[0].lower()
    tweet += " " + URGENT_LEVEL[urgent]+":"

####    # Need Brigadists
####    needHelp = (row[1].lower() == "si") or (row[1].lower() == "sí")
####    if (needHelp):
####        tweet += ", NECESITAN BRIGADISTAS"
##
##    
    # Address
    address = row[5]
    tweet += " Rescate en " + address
    
    # Zone
    zone = row[6]
    tweet += " " + zone
        
    print(tweet)

#UrgentLevel (X)
#NEED BRIGADISTS (X)
#MOST IMPORTANT REQUIRED
#ADMITTED
#NOT REQUIRED
#ADDRESS
#ZONE
#DETAIL/SOURCE
#UPGRADE

main()
