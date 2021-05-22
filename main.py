from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient import discovery
import numpy as np
import itertools
from pprint import pprint
import pandas as pd


# If modifying these scopes, delete the file token.pickle.
scopes = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
# spreadsheet_id = '1NiUwgPP1evcs7WkXRlr2vnyvB_cXpswnKrrebrnHl6w'
spreadsheet_id = '1XMp85fU0Gh7i_sawUHnxw3P_4iQK4HmnL4iXY8rF8wI'


def main():

    print("Setting up...")


    """
        CREDENTIALS AND SETUP
    """


    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)


    """
        VARIABLES
    """


    # get the data from the sheet
    sheet = service.spreadsheets()
    vars = sheet.values().get(spreadsheetId=spreadsheet_id, range="Vars").execute().get('values', [])
    vars = list(map(list, itertools.zip_longest(*vars, fillvalue=None))) # transpose list
    votes_list = sheet.values().get(spreadsheetId=spreadsheet_id, range="Ranked-Choice!K3:T33").execute().get('values', [])

    votes_list = [[x or '0' for x in xs] for xs in votes_list] # replace empty cells with zeros
    votes_list = [[int(x) for x in xs] for xs in votes_list] # convert everything to a number

    # convert to a pandas df and back again to add in trailing columns/rows
    df = pd.DataFrame(votes_list)
    df_replace = df.fillna(0) # replace empty cells with 0s
    votes_list = df_replace.values.tolist()

    votes_list_by_ppl = list(map(list, itertools.zip_longest(*votes_list, fillvalue=0))) # transpose list

    num_books_to_pick = int(vars[2][1])

    # get the names of the people voting
    ppl = vars[1]
    del ppl[0]
    ppl = [p for p in ppl if p] # remove none values

    # get the list of books being voted for
    books = vars[0]
    del books[0]
    books = [b for b in books if b] # remove none values

    # set up the dictionary (by book title) for the votes
    votes_dict = {}
    for i in range(len(books)):
        votes_dict[books[i]] = votes_list[i]

    # set up the dictionary (by book title) for tiebreaking
    tiebreak = {}
    for i in range(len(books)):
        sum = 0
        for j in range(len(votes_list[i])):
            if votes_list[i][j] == 0:
                sum += 30
            else:
                sum += votes_list[i][j]
        tiebreak[books[i]] = sum

    # set up dictionary (by people) for the votes
    votes_dict_by_ppl = {}
    for i in range(len(ppl)):
        votes_dict_by_ppl[ppl[i]] = votes_list_by_ppl[i]

    # set up dictionary (by book) for people who have their vote for that book
    ppl_per_book = {}
    for b in books:
        ppl_per_book[b] = {}

    # current score of each book
    scores = {}
    for b in books:
        scores[b] = 0


    """
        DO THE RANKED-CHOICE PROCESS
    """


    # 0th round
    round_scores = []
    for key, value in votes_dict.items():
        for i in range(len(value)):
            if value[i] == 1:
                scores[key] += 1;
                ppl_per_book[key][ppl[i]] = 1
    round_scores.append(scores.copy())


    # the rest of the rounds

    col = 23
    alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ']
    round_num = 1

    while True:

        print("Computing round " + str(round_num) + "...")

        # if a majority on one book has been reached
        max_score = max(scores.values())
        if max_score > (len(ppl)/2):
            break

        # find book(s) with lowest score(s) that are not 0
        val = min(filter(None, scores.values()))
        res = [k for k, v in scores.items() if v == val]

        # tiebreak if necessary (we want to eliminate the largest total score)
        tb_score = -1
        tb_book = "";
        if len(res) > 1:
            for r in res:
                if tiebreak[r] > tb_score:
                    tb_score = tiebreak[r]
                    tb_book = r
            elim = tb_book
        else:
            elim = res[0]

        # reassign votes for people who have their vote for that book
        for p, v in ppl_per_book[elim].items():

            max_vote = max(votes_dict_by_ppl[p]) # get the max of the votes which that person voted for

            v += 1 # get next-place vote
            if v > max_vote: break
            index = votes_dict_by_ppl[p].index(v)
            to_add = books[index]

            while(scores[to_add] == 0): # make sure that the book is still in the running
                v += 1
                if v > max_vote: break
                index = votes_dict_by_ppl[p].index(v)
                to_add = books[index]

            if v <= max_vote: # make sure we don't count votes for ppl with no votes for books left in the running
                scores[to_add] += 1
                ppl_per_book[to_add][p] = v

        ppl_per_book[elim] = {}
        scores[elim] = 0

        # update the sheet
        scores_to_sheet = []
        for b, s, in scores.items():
            if s == 0:
                scores_to_sheet.append("")
            else:
                scores_to_sheet.append(s)

        service = discovery.build('sheets', 'v4', credentials=creds)
        value_input_option = 'RAW'
        sheet_range = "Ranked-Choice!" + alphabet[col] + "3"
        value_range_body = {
            "range": sheet_range,
            "majorDimension": "COLUMNS",
            "values": [
                scores_to_sheet
            ]
        }

        request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=sheet_range, valueInputOption=value_input_option, body=value_range_body)
        response = request.execute()

        round_scores.append(scores.copy())
        col += 1
        round_num += 1


    # final ranking

    print("Ranking top " + str(num_books_to_pick) + " books...")

    ranked = []
    for i in range(len(books)):
        ranked.append(0)

    rank = 1
    rank_col = []
    for k in range(len(books)):
        rank_col.append("")

    for i in range(len(round_scores)):

        if num_books_to_pick == 0: # if we have picked the necessary number of books
            break

        # get book
        j = len(round_scores) - i - 1
        round = round_scores[j]
        max_score = 0
        to_rank = ""
        for b, s in round.items():
            if s > max_score and ranked[books.index(b)] == 0:
                max_score = s
                to_rank = b

        # rank book
        row = books.index(to_rank)
        ranked[row] = 1
        rank_col[row] = rank

        rank += 1
        num_books_to_pick -= 1


    # if we need to pick more books
    for i in range(num_books_to_pick):
        min_score = -1
        to_rank = ""
        for b, s in tiebreak.items():
            if (s < min_score or min_score == -1) and ranked[books.index(b)] == 0:
                min_score = s
                to_rank = b

        # rank book
        row = books.index(to_rank)
        ranked[row] = 1
        rank_col[row] = rank

        rank += 1


    # update google sheet
    service = discovery.build('sheets', 'v4', credentials=creds)
    value_input_option = 'RAW'
    sheet_range = "Ranked-Choice!U3"
    value_range_body = {
        "range": sheet_range,
        "majorDimension": "COLUMNS",
        "values": [
            rank_col
        ]
    }

    request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=sheet_range, valueInputOption=value_input_option, body=value_range_body)
    response = request.execute()



if __name__ == '__main__':
    main()
