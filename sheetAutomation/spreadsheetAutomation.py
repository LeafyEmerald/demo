# to make the spreadsheets work
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# to make the reddit side work
import praw
import json


# THE SHEETS PART
scope = ['https://spreadsheets.google.com/feeds',
'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('Project Katsuki-42ad85cb5253.json', scope)

gc = gspread.authorize(credentials)

sh = gc.open('Conservative_Whip Sheet_FOR_TESTING')

#determines what tab/worksheet we're doing the whole searcing of the data from
worksheet = sh.get_worksheet(1)

emptyRowFound = False
# c = column, we start at one as there's no collum 0
c = 1

# loop to check for empty collumns, we'll need to eventually take another look at this
# to add certain exceptions that are needed for this to play ball with the actual whip
# sheet.
# the name could be anyone technically
while emptyRowFound != True:
    
    values_list = worksheet.col_values(c)

    if 'AYE' in values_list or 'NO' in values_list or 'ABS' in values_list or 'InfernoPlato' in values_list or 'Whip responsible' in values_list or 'Whip region' in values_list or 'vote' in values_list:
        c += 1
    
    else:
        emptyRowFound = True   
        print(c)

# a loop to fill up a row with values

## reddit implementation

with open('config.json', 'r') as myfile:
    data = myfile.read()
    myfile.close()

obj = json.loads(data)

# does the whoel Oauth thing, uses the info we have from the config.json file

reddit = praw.Reddit(client_id = str(obj['client_id']), client_secret = str(obj['client_secret']), user_agent = str(obj['user_agent']))

# user input

urlFound = False

while urlFound != True:
    
    print("What bill would you like to count?")
    url = input()

    try: 
        submission = reddit.submission(url=url)
        urlFound = True
    except:
        print("That wasn't a correct URL. Please verify that you have typed the URL correctly and try again.")

#creates array of both names and votes
names, votes=[], []

for top_level_comment in submission.comments:
    if (top_level_comment.author_flair_css_class == 'conservative'):
        # dynamically creates list of names and votes
        names.append(top_level_comment.author.name)

        # makes it all neatly into capitalised letters
        vote = top_level_comment.body
        vote = vote.upper()
        votes.append(vote)

names_and_votes = [{'name': n, "vote": v} for n, v in zip(names,votes)]

with open('output.json', 'w') as outfile:
    json.dump(names_and_votes, outfile)
    print("done!!")
    outfile.close()

# reads content of the file

with open('output.json', 'r') as input:
    data = input.read()

obj = json.loads(data)

for element in obj:
    # we're gonna use this to find which row they're on
    try:
        cell = worksheet.find(element["name"])
        row_number = cell.row
        print(row_number)
        worksheet.update_cell(row_number, c, element["vote"])
        # print("mp has been founded")

    except:
        print("mp hasn't voted, or is not found on the spreadsheet")       
