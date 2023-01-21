import pandas as pd
import numpy as np
import openpyxl
import re

# Steps:
# 1 - Download the google sheet as xlsx. Save to this directory as sheet.xlsx
# 2 - Run generate-loot-list.py
# 3 - Review output LootTable.lua in this directory. If it looks good, copy it to the root of the repo 

# Ignore these sheets from the spreadhseet
IGNORE_SHEETS = ['Introduction', 'Physical BIS Lists', 'Caster BIS Lists', 'Healer BIS Lists', 'Tank BIS Lists']
FILENAME = "sheet.xlsx"
OUTPUT_FILENAME = "LootTable.lua"

SHEET_SPECS = {
    "Physical Loot": {
        "ignore_rows_before": 1, # First two rows
        "item_name_col": 3, # Col: D
        "prio_col": 17, # Col: R
        "notes_col": 18 # Col: S
    },
    "CasterHealer Loot": {
        "ignore_rows_before": 1, # First two rows
        "item_name_col": 3, # Col: D
        "prio_col": 15, # Col: P
        "notes_col": 16 # Col: Q
    },
    "Tank Loot": {
        "ignore_rows_before": 1, # First two rows
        "item_name_col": 2, # Col: C
        "prio_col": 16, # Col: Q
        "notes_col": 17 # Col: R
    }
}
# It seems that not all tabs (sheets) are created equal. Some have boss columns, some don't.

#  https://gist.github.com/zachschillaci27/887c272cbdda63316d9a02a770d15040
def _get_link_if_exists(cell):
    try:
        return cell.hyperlink.target
    except AttributeError:
        return None

def _get_item_id_from_link(link):
    # We expect a link that contains /item=####/
    if link is None:
        return None

    matcher = re.match(r"\S+\/item=(\d+)", link)
    if matcher:
        return matcher.groups()[0]
    return None


# Main logic...

doc = pd.ExcelFile(FILENAME)
lootTable = {} # Key: ItemID, Value: Dict
for sheetName in doc.sheet_names:
    print("Processing sheet: {}".format(sheetName))

    if sheetName in IGNORE_SHEETS:
        print(">> Ignoring")
        continue

    sheet = pd.read_excel(doc, sheetName)
    sheet = sheet.replace({np.NaN:None})
    openpysheet = openpyxl.load_workbook(FILENAME)[sheetName]

    if sheetName not in SHEET_SPECS:
        print("ERROR: No sheet spec for sheet. Skipping")
        continue

    spec = SHEET_SPECS[sheetName]
    ignoreRowsBeforeIndex = spec["ignore_rows_before"]
    itemNameColIndex = spec["item_name_col"]
    prioColIndex = spec["prio_col"]
    notesColIndex = spec["notes_col"]

    for index, row in sheet[ignoreRowsBeforeIndex:].iterrows():
        # each row is returned as a pandas series
        print("Row: {}".format(index))

        itemName = row[itemNameColIndex]
        itemLink = _get_link_if_exists(openpysheet.cell(row=index+ignoreRowsBeforeIndex+1, column=itemNameColIndex+1))
        itemId = _get_item_id_from_link(itemLink)
        prioText = row[prioColIndex]
        notesText = row[notesColIndex]

        print("Item Name: {}".format(itemName))
        print("Item ID: {}".format(itemId))
        print("Prio: {}".format(prioText))
        print("Notes: {}".format(notesText))
        print("--------------------------")

        if itemId is None:
            print("ERROR: Unable to extract item id. Skipping.")
            continue

        if itemName is None:
            print("ERROR: Unable to extract item name. Skipping.")
            continue

        if prioText is None:
            print("WARNING: No prio text for item. Skipping.")
            continue

        lootSheetEntry = {
            "sheet": sheetName,
            "prio": prioText
        }

        if notesText is not None:
            lootSheetEntry["note"] = notesText

        lootEntry = {
            "itemid": itemId,
            "itemname": itemName,
            "sheets": [
                lootSheetEntry
            ]
        }

        if itemId not in lootTable:
            # New item
            lootTable[itemId] = lootEntry
        else:
            # Already exists in another sheet
            lootTable[itemId]["sheets"].append(lootEntry["sheets"][0])

print("Writing loot table to {filename}".format(filename=OUTPUT_FILENAME))

with open(OUTPUT_FILENAME, "w") as f:
    f.write("lootTable = {\n")

    # { ["itemid"] = "45110", ["itemname"] = "Titanguard", ["sections"] = {{["sheet"] = "Physical DPS", ["prio"] = "Tank", ["note"] = "Give to your MT first"},{["sheet"] = "Caster DPS", ["prio"] = "Blah", ["note"] = "Howdy"}}},


    # TODO: Support multi sheet
    for key, value in lootTable.items():

        sectionsText = "{"
        for sheet in value["sheets"]:
            if "note" in sheet:
                sectionsText = sectionsText + "{{[\"sheet\"] = \"{sheetname}\", [\"prio\"] = \"{prio}\", [\"note\"] = \"{note}\"}},".format(
                    sheetname=sheet["sheet"],
                    prio=sheet["prio"],
                    note=sheet["note"]
                )
            else:
                sectionsText = sectionsText + "{{[\"sheet\"] = \"{sheetname}\", [\"prio\"] = \"{prio}\"}},".format(
                    sheetname=sheet["sheet"],
                    prio=sheet["prio"]
                )
        sectionsText = sectionsText + "}"

        outputLine = "{{[\"itemid\"] = \"{itemid}\", [\"itemname\"] = \"{itemname}\", [\"sections\"] = {sections}}},\n".format(
            itemid=value["itemid"],
            itemname=value["itemname"],
            sections=sectionsText
        )
        f.write(outputLine)

    f.write("}\n")

print("Done")