import valo_api
import gspread
from datetime import datetime
import re
import time
import xlsxwriter

from urllib.parse import unquote


sa = gspread.service_account(filename='credentials.json')
sh = sa.open("HoojNotes")
wks = sh.worksheet("Students")
wks2 = sh.worksheet("Notes")
valoapikey = ""
valo_api.set_api_key()

def getID(url):
    name = re.search(r'riot/(.*?)(?=%23)', url).group(1)
    tag = re.search(r'%23(.*?)/', url).group(1)
    name = unquote(name)
    return [name, tag]

def isOld(date_string):
    # Convert input string to datetime object
    date = datetime.strptime(date_string, "%m/%d/%Y")

    # Define the comparison date
    comparison_date = datetime(2023, 5, 15)

    # Compare the dates
    if date < comparison_date:
        return True
    else:
        return False
    
# def isOlder(dates1, dates2) :
#     date1 = datetime.strptime(dates1, "%m/%d/%Y")
#     date2 = datetime.strptime(dates2, "%m/%d/%Y")
    
#     if date1 < date2:
#         return True
#     else:
#         return False
    
def latestEA(dates):
    date = datetime.strptime(dates, "%m/%d/%Y")
    e5a2 = datetime.strptime("10/17/2022", "%m/%d/%Y")
    e5a3 = datetime.strptime("1/9/2023", "%m/%d/%Y")
    e6a1 = datetime.strptime("3/6/2023", "%m/%d/%Y")
    e6a2 = datetime.strptime("6/10/2023", "%m/%d/%Y")  
    
    if date < e5a2:
        return 0
    elif date < e5a3:
        return 1
    elif date < e6a1:
        return  2
    elif date < e6a2:
        return 3
    else:
        return 4
        
acts = ["e5a2", "e5a3", "e6a1", "e6a2", "e6a3"]      
    
    
       
def is_higher_rank(rank1, rank2):
    if rank2 == "Radiant":
        return True
    if rank2 == "Unrated":
        return False
    tiers = ["Iron", "Bronze", "Silver", "Gold", "Platinum", "Diamond", "Ascendant", "Immortal"]
    divisions = [1, 2, 3]

    # Split the rank strings into tier and division
    rank1_tier, rank1_division = rank1.split()
    rank2_tier, rank2_division = rank2.split()

    # Compare the tiers
    if tiers.index(rank2_tier) > tiers.index(rank1_tier):
        return True
    elif tiers.index(rank2_tier) < tiers.index(rank1_tier):
        return False

    # If tiers are equal, compare the divisions
    if int(rank2_division) > int(rank1_division):
        return True
    elif int(rank2_division) < int(rank1_division):
        return False

    # If both tiers and divisions are equal, ranks are the same
    return False
    
    
results = []
missing = []


# print(mm.highest_rank)
# for y in mm.by_season:
#     x = mm.by_season[y]
#     print(y, x.final_rank, x.final_rank_patched, x.old)
#     if x.act_rank_wins is not None:
#         for x in x.act_rank_wins:
#             print(x)




s1 = wks.get_all_records()
s2 = wks2.get_all_records()
idx = 0

for x in s1:
    print(x)
    date = s2[idx]["Date"]
    # print(x["Name"], x["Tracker"])
    if(isOld(date) and x["Tracker"] != "" and ("https://tracker.gg/valorant/profile/riot/" in x["Tracker"])):
        riottag = getID(x["Tracker"])
        exists = True
        region = "na"
        eu = True
        try:
            mmregion = valo_api.get_match_history_by_name_v3("eu", riottag[0], riottag[1], 1)
        except:
            eu = False
        if eu:
            region = "eu"
            
        ap = True
        try:
            mmregion = valo_api.get_match_history_by_name_v3("ap", riottag[0], riottag[1], 1)
        except:
            ap = False
        if ap:
            region = "ap"
        try:
            mm = valo_api.get_mmr_details_by_name_v2(region, riottag[0], riottag[1])
        except:
            exists = False
        highest = x["Starting Rank"]
        if "Immortal" in highest:
            highest = highest[0:10]
        if "Radiant" not in highest:
            if exists:
                ea = latestEA(date)
                for i in range(ea, 5):
                    epiacts = mm.by_season[acts[i]]
                    if epiacts.act_rank_wins is not None:
                        if is_higher_rank(highest, epiacts.act_rank_wins[0].patched_tier):
                            highest = epiacts.act_rank_wins[0].patched_tier
                results.append([x["Name"], x["Starting Rank"], highest, x["Tracker"], date])
            else:
                missing.append([x["Name"], x["Starting Rank"], x["Tracker"], date])
    else:
        missing.append([x["Name"], x["Starting Rank"], x["Tracker"], date])    
    while idx < len(s2) - 1 and s2[idx]["Name"] == x["Name"]:
            idx += 1
            #print(idx, len(s2))
    
    
workbook = xlsxwriter.Workbook('whj.xlsx')
worksheet = workbook.add_worksheet()

row = 0

for col, data in enumerate(results):
    worksheet.write_column(row, col, data)
    
workbook.close()

workbook = xlsxwriter.Workbook('missing.xlsx')
worksheet = workbook.add_worksheet()

row = 0

for col, data in enumerate(missing):
    worksheet.write_column(row, col, data)

workbook.close()