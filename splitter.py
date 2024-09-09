import pandas as pd
import re
import time


# ----------------- HELPERS -------------------


# row number then column number. takes cell x,y and returns the value
def returnCellValue(x,y):
    return df.iloc[x, y]


# returns a list of cell values. Takes x which is row number.
def returnRow(x):
    return df.iloc[x, 0:]


# returns a list of cell values from 1 column.
def returnCol(y):
    return df.iloc[0:, y]


# given a key, returns specific value in the dictionary.
# if not found, create the new dataframe and return it.
def dictHelper(name):
    if name in hdict:
        temp = hdict.get(name)
        return temp
    else:
        hdict.update({name: pd.DataFrame([], columns=df.columns)})
        return hdict.get(name)

def dictUpdater(name, df):
    hdict.update({name: df})

# ---------------------------------------------

df = pd.read_excel('2024 Fall Team Application (Responses).xlsx')
# print(df)

hdict = {}

# Column # of data that we are filtering by.
filterIndex = ord(input('Enter column letter to filter games by. Case does not matter: ').lower()) - 97

startTime = time.time()

# Iterate through the entire sheet. Main logic of where it happens
for index, row in df.iterrows():
    games = returnCellValue(index, filterIndex)
    game_lst = games.split(', ')
    for game in game_lst:
        prevDf = dictHelper(game)
        append = pd.DataFrame([row.values], columns=df.columns)
        if prevDf.shape[0] == 0:
            dictUpdater(game, append)
        else:
            newDf = pd.concat([dictHelper(game), append])
            dictUpdater(game, newDf)

writer = pd.ExcelWriter("SortedTeams.xlsx")

for k, v in hdict.items():
    newKey = re.sub('[^A-Za-z0-9 ]+', "", k)
    v.to_excel(writer, sheet_name=newKey, index=False)
    # Automatic spacing
    worksheet = writer.sheets[newKey]  # pull worksheet object
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = series.astype(str).map(len).max() + 1  # length of maximum item + 1 for extra space
        worksheet.set_column(idx, idx, max_len)  # set column width

writer.close()

endTime = time.time()

print(f"Finished in {endTime-startTime} seconds")
