import gspread
import json
import praw
import re
import time
from gspread.exceptions import APIError
from oauth2client.service_account import ServiceAccountCredentials

def main():
    scope = ['https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive']
    
    gc = gspread.authorize(authenticateSheetsAPI(scope))
    # we will eventually rework this into user input, same with the worksheet being used as well -- will consider even making into a seperare function
    sh = gc.open('Conservative_Whip_Sheet_FOR_TESTING')
    worksheet = sh.get_worksheet(2)

    # we check what the readConfig function returns
    if readConfig() == True:
        with open('config.json', 'r') as f:
            data = json.load(f)  
        lookupRedditURL(authenticateRedditAPI(), worksheet, str(data['lastEmptyCollumn']))
    
    else:
        with open('config.json', 'r') as f:
            data = json.load(f)
            data['lastEmptyCollumn'] = findEmptyCollumn(worksheet)
                
        with open ('config.json', 'w') as f:
            json.dump(data, f)
        
        lookupRedditURL(authenticateRedditAPI(), worksheet, str(data['lastEmptyCollumn']))

def authenticateSheetsAPI(scope):
    credentials = ServiceAccountCredentials.from_json_keyfile_name('Project Katsuki-42ad85cb5253.json', scope)
    return credentials

def readConfig():
    with open('config.json', 'r') as f:
        data = f.read()
        objectData = json.loads(data)

    try:
        print(str(objectData['lastEmptyCollumn']))
        return True
    except:
        return False

def findEmptyCollumn(worksheet, c = 1):
    emptyCollumnFound = False
    timerStarted = False

    while emptyCollumnFound != True:
        try:
            while timerStarted != True:
                start_time = time.time()
                timerStarted = True

            collumn_values = worksheet.col_values(c)

            if 'AYE' in collumn_values or 'NO' in collumn_values or 'ABS' in collumn_values or 'InfernoPlato' in collumn_values or 'Whip Responsible' in collumn_values or 'Whip Region' in collumn_values or 'Attendance' in collumn_values or 'Total' in collumn_values or '%' in collumn_values or 'Swear-in' in collumn_values: 
                c += 1
            else: 
                emptyCollumnFound = True
                return(c)

        except APIError:
            timerStarted = False
            print("this will take a moment!")
            time.sleep(100 - (time.time() - start_time))

def authenticateRedditAPI():
    with open('config.json', 'r') as f:
        obj = json.load(f) 
    return praw.Reddit(client_id = str(obj['client_id']), client_secret = str(obj['client_secret']), user_agent = str(obj['user_agent']))

def lookupRedditURL(reddit, worksheet, collumn):
    urlFound = False

    while urlFound != True:
        print("What Bill would you like to count?")
        url = input()

        try: 
            submission = reddit.submission(url=url)
            urlFound = True
            lookupRedditComments(reddit, worksheet, collumn, submission)

        except:
            print("That wasn't a correct Reddit link. Please verify that it is a correct link and try again")



def lookupRedditComments(reddit, worksheet, collumn, submission):
    
    names, votes = [], []
    nameList = lookupNames(worksheet)
    # print(nameList)
    # number of MPs

    number_of_mps = len(nameList)

        #creates array of both names and votes
    for top_level_comment in submission.comments:
        if (top_level_comment.author_flair_css_class == 'conservative'):
            vote = top_level_comment.body
            if re.search('proxy', vote, re.IGNORECASE):
                split_proxy_vote = vote.split()
                    
                i = 0

                while i < number_of_mps:
                    for word in split_proxy_vote:
                        if nameList[i] in word:
                            names.append(nameList[i])
                            # break
                    i += 1

                for word in split_proxy_vote:
                    if re.search('aye', word, re.IGNORECASE) or re.search('no', word, re.IGNORECASE) or re.search('abstain', word, re.IGNORECASE):
                        word = word.upper()
                        if word == 'ABSTAIN':
                            votes.append(word[:3])
                        else:
                            votes.append(word)
                
            elif re.search('aye', vote, re.IGNORECASE) or re.search('no', vote, re.IGNORECASE) or re.search('abstain', vote, re.IGNORECASE):
                names.append(top_level_comment.author.name)
                votes.append(vote.upper())

    names_votes = [{'name': n, "vote": v} for n, v in zip(names,votes)]

    with open('output.json', 'w') as outfile:
        json.dump(names_votes, outfile)
        outfile.close()

    writeToSheet(worksheet, collumn)

def lookupNames(worksheet):
    # note for future self: this here will need some fixing/optimisation
    nameList = worksheet.col_values(3)
    del nameList[:3]
    return nameList

def writeToSheet(worksheet, collumn):

    with open('output.json', 'r') as input:
        data = input.read()

        obj = json.loads(data)

    for element in obj:
        try:
            # eventually check if we can regex this
            name = re.compile(element["name"], re.IGNORECASE)
            cell = worksheet.find(name)
            worksheet.update_cell(cell.row, collumn, element["vote"])
        except:
            print("That MP wasn't found on the sheet.")
    
    with open('config.json', 'r') as f:
        data = f.read()
        obj = json.loads(data)
        obj['lastEmptyCollumn'] = int(obj['lastEmptyCollumn']) + 1

    with open('config.json', 'w') as f:
        json.dump(obj, f)

if __name__ == "__main__":
    main()
