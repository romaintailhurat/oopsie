"""
Main logic for producing the calendar.
"""
from json import load
from re import match
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook



def months_years():
    months = ["janvier", "février", "mars", "avril", "mai","juin", "juillet", "aout", "septembre", "octobre", "novembre", "decembre"]
    base_cols = [m + " 2023" for m in months] # TODO for every year

    base_cols = []

    for year in ["2023", "2024", "2025"]:
        base_cols += [m + " " + year for m in months]

    return(base_cols)

def init_calendar(my):
    cols = { "date" : my}
    return(pd.DataFrame.from_dict(cols))

def produce_calendar(trello_json):

        trello = load(trello_json)
        columns = {lst["id"]:lst["name"] for lst in trello["lists"]}    
        my = months_years()
        calendar = init_calendar(my)

        for card in trello["cards"]:
            # 2023, 2024, 
            if card["idList"] in ["63242327b8fb6200c269a8b2", "63242338083bc40060dedc05", "63247cfd74c0f3049c8eff0e"]:
                rule = r"\w+ : \w+ [0-9]{4}"
                rule_match = match(rule, card["name"])

                if rule_match is not None:
                    i, j = rule_match.span()
                    operation = card["name"][i:j]
                    name, date = operation.split(":")
                    date = date.strip()                

                    try:
                        i = my.index(date)
                        # Column already created
                        if name in calendar.columns:
                            calendar.at[i, name] = " - ".join([label["name"] for label in card["labels"]])
                        # Column has to be created
                        else:
                            col = ["" for _ in range(36)] # FIXME 12 * 3, Magic number, LOL                
                            col[i] = " - ".join([label["name"] for label in card["labels"]])
                            calendar[name] = col
                    except ValueError:
                        print(f"Bad date format with {date}")

        wb = Workbook()
        ws = wb.active

        for row in dataframe_to_rows(calendar.T, index=True, header=True):
            ws.append(row)

        #wb.save("C:/Users/ARKN1Q/Downloads/calendrier.xslx")
        return(wb)