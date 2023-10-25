import pandas as pd
import numpy as np
import openpyxl
import json
import re

# Steps:
# 1 - Download the google sheet as xlsx. Save to this directory as sheet.xlsx
# 2 - Run generate-loot-list.py
# 3 - Review output LootTable.lua in this directory. If it looks good, copy it to the root of the repo 

# Ignore these sheets from the spreadhseet
IGNORE_SHEETS = ['Introduction', 'Physical BIS Lists', 'Caster BIS Lists', 'Healer BIS Lists', 'Tank BIS Lists', 'Fusion Streams']
FILENAME = "joardee-phase-four.xlsx"
OUTPUT_FILENAME = "PhaseFourLootTable.lua"
errors = []
warnings = []
# Fix for typos in the spreadsheet
nameoverrides = {
    "Byrntroll, the Bone Arbiter": "Bryntroll, the Bone Arbiter",
    "Bulwark of Smoldering Steel": "Bulwark of Smouldering Steel",
    "Leather Stitched Scourge Parts": "Leather of Stitched Scourge Parts",
    "Bloodsurge, Kel'thuzad's Blade of Agony": "Bloodsurge, Kel'Thuzad's Blade of Agony",
    "Devium's Eternall Cold Ring": "Devium's Eternally Cold Ring"
}

SHEET_SPECS = {
    "Physical Loot": {
        "alias": "Physical Loot",
        "ignore_rows_before": 1, # First two rows
        "item_name_col": 2, # Col: C
        "prio_col": 16, # Col: Q
        "notes_col": 17 # Col: R
    },
    "CasterHealer Loot": {
        "alias": "Caster/Healer Loot",
        "ignore_rows_before": 1, # First two rows
        "item_name_col": 2, # Col: C
        "prio_col": 14, # Col: O
        "notes_col": 15 # Col: P
    },
    "Tank Loot": {
        "alias": "Tank Loot",
        "ignore_rows_before": 1, # First two rows
        "item_name_col": 2, # Col: C
        "prio_col": 14, # Col: O
        "notes_col": 15 # Col: P
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

def _get_item_ids_from_json_loot_table(name):
    """ Attempts to get the item ids from the backup-loot-table.json file"""

    # Read backup-loot-table.json as json
    # Attempt to find the name in the json dict
    # Return the item id(s) if found
    # Return None if not found

    with open("backup-loot-table.json", "r") as f:
        lootTable = json.load(f)
        if name in lootTable:
            return lootTable[name]
        else:
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
    sheetAlias = spec["alias"]

    for index, row in sheet[ignoreRowsBeforeIndex:].iterrows():
        # each row is returned as a pandas series
        print("Row: {}".format(index))

        itemName = row[itemNameColIndex]
        itemLink = _get_link_if_exists(openpysheet.cell(row=index+ignoreRowsBeforeIndex+1, column=itemNameColIndex+1))

        # Sometimes there are typos in the sheet. We can override the name here
        if itemName in nameoverrides:
            print(">> Overriding item name: {}".format(nameoverrides[itemName]))
            itemName = nameoverrides[itemName]

        if itemLink is not None:
            print("Item link: {}".format(itemLink))
            itemId = _get_item_id_from_link(itemLink)
        else:
            print("Attempting to get item id from backup-loot-table.json")
            itemId = _get_item_ids_from_json_loot_table(itemName)
        prioText = row[prioColIndex]
        notesText = row[notesColIndex]

        # If notesText is a string, remove all newline characters
        if isinstance(notesText, str):
            notesText = notesText.replace("\n", "")

        print("Item Name: {}".format(itemName))
        print("Item ID: {}".format(itemId))
        print("Prio: {}".format(prioText))
        print("Notes: {}".format(notesText))
        print("--------------------------")

        if itemName is None:
            print(f"ERROR: Unable to extract item name (row: {row}, index: {index}). Skipping.")
            errors.append(f"ERROR: Unable to extract item name (row: {row}, index: {index}). Skipping.")
            continue

        if itemId is None:
            print(f"ERROR: Unable to extract item id for {itemName}. Skipping.")
            errors.append(f"ERROR: Unable to extract item id for {itemName}. Skipping.")
            continue

        if prioText is None:
            print(f"WARNING: No prio text for item {itemName}. Skipping.")
            warnings.append(f"WARNING: No prio text for item {itemName}. Skipping.")
            continue

        lootSheetEntry = {
            "sheet": sheetAlias,
            "prio": prioText
        }

        if notesText is not None:
            lootSheetEntry["note"] = notesText

        # Check if itemId is a list. If so, we need to add multiple entries to the loot table.
        # This happens when we used the backup-loot-table.json file to get the item id
        if isinstance(itemId, list):
            # itemId is a list
            for itemIdEntry in itemId:
                lootEntry = {
                    "itemid": itemIdEntry,
                    "itemname": itemName,
                    "sheets": [
                        lootSheetEntry
                    ]
                }

                if itemIdEntry not in lootTable:
                    # New item
                    lootTable[itemIdEntry] = lootEntry
                else:
                    # Already exists in another sheet
                    lootTable[itemIdEntry]["sheets"].append(lootEntry["sheets"][0])
        else:
            # itemId is a single value
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

    # Example item line
    # { ["itemid"] = "45110", ["itemname"] = "Titanguard", ["sections"] = {{["sheet"] = "Physical DPS", ["prio"] = "Tank", ["note"] = "Give to your MT first"},{["sheet"] = "Caster DPS", ["prio"] = "Blah", ["note"] = "Howdy"}}},

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


if len(warnings) > 0:
    print("Warnings encountered:")
    for warning in warnings:
        print(warning)

if len(errors) > 0:
    print("Errors encountered:")
    for error in errors:
        print(error)
