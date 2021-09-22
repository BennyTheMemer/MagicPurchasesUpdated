from bs4 import BeautifulSoup as bs
from selenium import webdriver
import time
import pandas as pd
import undetected_chromedriver as uc
import requests
import gspread
from gspread_formatting import *


gc = gspread.service_account(filename="keys.json")
sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1f65bujzYwsMcYvAgxXP-NRlUSTGt0dkRhcln63Cpar8/edit#gid=0").sheet1

def main():

    values = sh.get_all_values()
    head = values.pop(0)
    list_of_dicts = [{head[i]: col for i, col in enumerate(row)} for row in values]

    for n in list_of_dicts:
        url =n["Link"]
        response =requests.get(url)
        soup = bs(response.content, 'html.parser')

        # If the product doesn't have "reprints" len of the tag will be 8
        if len(soup.find_all("dt", {"class": "col-6 col-xl-5"})) == 8:
            trend = float(
                soup.find_all("dd", {"class": "col-6 col-xl-7"})[4].find("span").contents[0].strip(
                    " ").strip("€").replace(",", "."))
            minPrice = float(
                soup.find_all("dd", {"class": "col-6 col-xl-7"})[3].contents[0].strip(
                    " ").strip("€").replace(",", "."))
        # If the product does have "reprints" len of the tag will be 9
        else:
            trend = float(
                soup.find_all("dd", {"class": "col-6 col-xl-7"})[5].find("span").contents[0].strip(
                    " ").strip("€").replace(",", "."))
            minPrice = float(
                soup.find_all("dd", {"class": "col-6 col-xl-7"})[4].contents[0].strip(
                    " ").strip("€").replace(",", "."))
        n["Trend"] = trend
        WL = float(n["Trend"]) * int(n["Quantidade"]) - float(n["Comprada"]) * int(n["Quantidade"])
        WL = round(WL, 2)
        n["W/L"] = WL




    df = pd.DataFrame.from_dict(list_of_dicts)
    dataframe = pd.DataFrame.from_dict(df)
    sh.update([dataframe.columns.values.tolist()] + dataframe.values.tolist(), value_input_option="USER_ENTERED")
    rule = ConditionalFormatRule(
        ranges=[GridRange.from_a1_range('G1:G2000', sh)],
        booleanRule=BooleanRule(
            condition=BooleanCondition('NUMBER_GREATER', ['0']),
            format=CellFormat(backgroundColor=Color(0, 1, 0))
        )
    )
    rule2 = ConditionalFormatRule(
        ranges=[GridRange.from_a1_range('G1:G2000', sh)],
        booleanRule=BooleanRule(
            condition=BooleanCondition('NUMBER_LESS', ['0']),
            format=CellFormat(backgroundColor=Color(1, 0, 0))
        )
    )

    rules = get_conditional_format_rules(sh)
    rules.clear()
    rules.append(rule)
    rules.append(rule2)
    rules.save()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
