import requests
from bs4 import BeautifulSoup #to search and read from google
import pandas as pd #for excel data
import openpyxl #for excel
from pathlib import Path #to check if path is correct

latMultiplier = ""
longMultiplier = ""
latList = []
longList = []
skippedCodes = []

#checking if the input and output file exist
inputSatisfied = False
outputSatisfied = False
while(not inputSatisfied):
    inputSheet = input("Input: ")
    if(not Path(inputSheet).is_file()):
        print("File not found in directory. Try again")
    else:
        inputSatisfied = True

sheet = input("Sheet name: ")


while(not outputSatisfied):
    outputSheet = input("Output: ")
    if(not Path(outputSheet).is_file()):
        print("File not found in directory. Try again")
    else:
        outputSatisfied = True



#making a list of all the zip code values
a = pd.read_excel(inputSheet)
zipCodeList = a['ZIP'].values.tolist()
for b in range(len(zipCodeList)):
    zipCodeList[b] = str(zipCodeList[b])
    if(len(zipCodeList[b]) < 5):
        #pad with zeroes at the beginning if excel stored ZIPs as integers without
        #leading 0
        zipCodeList[b] = '0'*(5-len(zipCodeList[b]))+zipCodeList[b]

#running through the list, checking google, and adding the coordinates to the spreadsheet
for j in range(len(zipCodeList)):
    lat = ""
    longg = ""
    #use this link to google each ZIP code
    req = requests.get("https://www.google.com/search?q=zip+code+"+zipCodeList[j]+
    "+coordinates&rlz=1C1GCEA_enUS1022US1022&sxsrf=AJOqlzVUQ3XXB85tDLCOIKSAk-uFIr27wA%3A1674239465476&ei=6d3KY5DVHISzqtsP9OW_yAE&ved=0ahUKEwjQ7e6E5Nb8AhWEmWoFHfTyDxkQ4dUDCA8&uact=5&oq=zip+code+32709+coordinates&gs_lcp=Cgxnd3Mtd2l6LXNlcnAQAzIFCAAQogQyBQgAEKIEMgUIABCiBDIFCAAQogQyBQgAEKIEOgoIABBHENYEELADOggIIRCgARDDBDoECCMQJ0oECEEYAEoECEYYAFDSBFjTCWCMG2gBcAF4AIABe4gBzgOSAQMwLjSYAQCgAQHIAQjAAQE&sclient=gws-wiz-serp")
    soup = BeautifulSoup(req.content, "html.parser")
    result = soup.prettify()
    #look for this tag, which is typically found near the ZIP codes 
    if("BNeawe iBp4i AP7Wnd" in result):
        divIndex = result.index("BNeawe iBp4i AP7Wnd")
        
        #Seeing if a coordinate should be negative
        if(result[divIndex+113] == 'N'):
            latMultiplier = ""
        elif(result[divIndex+104] == 'S'):
            latMutliplier = "-"
        lat = latMultiplier+result[divIndex+104:divIndex+111]
        if(lat[len(lat)-1] == "°"):
            lat = lat[0:len(lat)-2]

        if(result[divIndex+125] == 'E'):
            longMultiplier = ""
        elif(result[divIndex+125] == 'W'):
            longMultiplier = "-"
        longg = longMultiplier+result[divIndex+116:divIndex+123]  
        if(longg[len(longg)-1] == "°"):
            longg = longg[0:len(longg)-2]

        
        latList.append(lat)
        longList.append(longg)
    #add them to a skippedCodes list if they are not found
    else:
        latList.append('Not Found')
        longList.append('Not Found')
        skippedCodes.append(zipCodeList[j])



print(skippedCodes)

#inserting the latitude and longitude columns into the original data
a.insert(2,"Lat", latList, True)
a.insert(3, "Long", longList, True)
#a = a.drop(['Unnamed: 0'], axis = 1)

#writing the updated data to the output sheet
a.to_excel(outputSheet)







