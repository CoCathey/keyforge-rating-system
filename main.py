# Required packages will be added here: 
# pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
# pip install gspread
# pip install oauth2client
# pip install fuzzywuzzy
# pip install python-Levenshtein

import time
import math
import gspread
from googleapiclient.discovery import build
from google.oauth2 import service_account
from oauth2client.service_account import ServiceAccountCredentials
from fuzzywuzzy import process

# Authentication stuff for gspread
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('keys.json', scope)
gc = gspread.authorize(credentials)
ranks_sheet = gc.open('Ranked Keyforge').sheet1
games_sheet = gc.open('Keyforge Games').sheet1
log = gc.open_by_key('1SCpcSCSjDYskm6zJ50twrNR5tYGGFoErWd0_pL3UCDE').sheet1
#####################################################

def spreadsheet():
# Getting into the Google Sheet
    SERVICE_ACCOUNT_FILE = 'keys.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None
    creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    SPREADSHEET_ID = '1FH9rSqHJMbtZ_ivDuGUW5Sw1xU-dw0kZjk4U69jgnkE'
    service = build('sheets', 'v4', credentials=creds)

#Getting the values from the ranks page of the sheet
    sheet = service.spreadsheets()
    s_result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
        range="ranks").execute()
    values=s_result.get('values',[])
    return values
#####################################################

def games_spreadsheet():
# Getting into the Google Sheet
    SERVICE_ACCOUNT_FILE = 'keys.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None
    creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    SPREADSHEET_ID = '1YRDXWeZPfABinj9zFhzBMc0-XuMufphT7xkuLjzDE7o'
    service = build('sheets', 'v4', credentials=creds)

#Getting the values from the ranks page of the sheet
    sheet = service.spreadsheets()
    s_result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
        range="Games").execute()
    values=s_result.get('values',[])
    return values
#####################################################

# A couple of global variables for the spreadsheet values
ranks_spreadsheet=spreadsheet()
games_spreadsheet=games_spreadsheet()
#####################################################

#Function to convert name from any input to the correct name in the sheet
def convertName(name):
    values = ranks_sheet.col_values(2)
    highestRatio = process.extractOne(name,values)
    return str(highestRatio[0])
#####################################################

#function to return player's rating
def rating(name,row):
    rating = ranks_sheet.cell(row, 3).value
    return int(rating)
##################################################

#function to update player's rating in spreadsheet
def ratingUpdate(name,rating,row):
    ranks_sheet.update_cell(row, 3, rating)
    return 0
##################################################

#function to return player's number of games
def games(name):
    row = playerRow(name)
    games = ranks_sheet.cell(row, 4).value
    return int(games)
##################################################

#function to return player's number of games
def wins(name,row):
    games = ranks_sheet.cell(row, 5).value
    return int(games)
##################################################

#function to return player's number of games
def losses(name,row):
    games = ranks_sheet.cell(row, 6).value
    return int(games)
##################################################

# Returns the row that the player's name is on
def playerRow(name):
    row = ranks_sheet.find(name).row
    return int(row)
##################################################

def rowValues(row):
    values = games_sheet.row_values(row)
    return values

#function to update winner's number of games in spreadsheet
def winUpdate(name,row):
    g_n=wins(name,row) + 1
    ranks_sheet.update_cell(row, 5, g_n)
    return 0
###################################################

#function to update loser's number of games in spreadsheet
def lossUpdate(name,row):
    g_n=losses(name,row) +1
    ranks_sheet.update_cell(row, 6, g_n)
    return 0
###################################################

# Function to determine win Probability
def Probability(rating1, rating2):
	return 1.0 * 1.0 / (1 + 1.0 * math.pow(10, 1.0 * (rating1 - rating2) / 400))
##################################################

# Function to calculate Elo rating
# k determines variability
def EloRatingWinner(Ra, Rb, k):
	# To calculate the Winning
	# Probability of Player B
	Pb = Probability(Ra, Rb)
	# To calculate the Winning
	# Probability of Player A
	Pa = Probability(Rb, Ra)
	Ra = Ra + k * (1 - Pa)
	Rb = Rb + k * (0 - Pb)
	return round(Ra)
    
def EloRatingLoser(Ra, Rb, k):
	# To calculate the Winning
	# Probability of Player B
	Pb = Probability(Ra, Rb)
	# To calculate the Winning
	# Probability of Player A
	Pa = Probability(Rb, Ra)
	Ra = Ra + k * (1 - Pa)
	Rb = Rb + k * (0 - Pb)
	return round(Rb)
##################################################

# Function to determine the k value 
def kvalue(name,w_sas,l_sas,keys_forged,gametype,row):
    g=games(name)
    r=rating(name,row)
    if r >= 2400:
        k=16
    elif r >= 2100:
        k=24    
    elif g <= 15:
        k=40
    else:
        k=32
    if w_sas > l_sas:
        k = k - ((w_sas-l_sas)/3)
    if l_sas > w_sas:
        k = k - ((w_sas-l_sas)/3)
    if keys_forged == 0:
        k = k+3
    if keys_forged == 2:
        k = k-3
    if gametype == "Tournament":
        k = k+3
    return k
##################################################

# Driver code below 

# Asks for the date of games you want to add to the rankings 
# It then finds those games and loops through driver code until
#   it finishes adding all the games 
date=input('enter date in format mm/dd/yyyy :\n')
values = games_spreadsheet
for line in values:
    if date in line[0]:
        game_row = games_sheet.find(line[0]).row
        row_Contents = rowValues(game_row)
        w = row_Contents[1]
        l = row_Contents[3]
        w_sas = row_Contents[2]
        l_sas = row_Contents[4]
        keys_forged = row_Contents[5]
        gametype = row_Contents[6]
        winner = convertName(w)
        loser = convertName(l)
        time.sleep(3)
        winner_row = playerRow(winner)
        loser_row = playerRow(loser)


# Finds the rating of each player 
        winner_rating = rating(winner,winner_row)
        loser_rating = rating(loser,loser_row)

# Finds the K value for each player based on given values 
        k_w = kvalue(winner,int(w_sas),int(l_sas),int(keys_forged),gametype,winner_row)
        k_l = kvalue(loser,int(w_sas),int(l_sas),int(keys_forged),gametype,loser_row)

# Finds the new rankings for each player 
        Ra=EloRatingWinner(winner_rating,loser_rating,k_w)
        Rb=EloRatingLoser(winner_rating,loser_rating,k_l)

# Updates the sheet with the new values 
        winUpdate(winner,winner_row)
        lossUpdate(loser,loser_row)
        ratingUpdate(winner,Ra,winner_row)
        ratingUpdate(loser,Rb,loser_row)

# Prints results to the terminal
        print("Previous Ratings:")
        print(winner, " : " ,winner_rating)
        print(loser, " : " ,loser_rating)
        print("New Ratings:")
        print(winner, " : ",Ra)
        print(loser, " : ",Rb)

        # This is a temporary delay because was getting an api overload error
        print('Temporary api delay to not overload quota')
        time.sleep(20)
        print('next game: ')


print("\nCOMPLETED!\n")


#### TO DO ####
# Rework Algorithm to do best rankings
# Would be cool to have option to type in name of 
#   deck instead of sas and it find the sas with the 
#   decks of keyforge api
# Maybe keep track of all the games it has calculated
#   Then checks with fuzzy wuzzy if it has calculated all the games
#   Then calculates any games not in the done list 
# Rank with the program and not with the rank function in sheets
#   That way people with less than 5 or 10 games can be considered unranked
# Any other suggestions? 


### BUGS 
# Added a delay as to not hit the API Quota 
#   Would like to get rid of this if possible
#   Trying to save rows for each game to minimize api calls
# Not really a bug, but need to find out the best k values and algorithm
