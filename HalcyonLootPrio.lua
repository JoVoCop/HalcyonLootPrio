-- Extracts a value from the loot table based on a provided key for a given itemLink
-- param: itemLink (string): format defined here: https://wowwiki-archive.fandom.com/wiki/API_GameTooltip_GetItem
-- param: key (string): "prio" or "note"
-- return: value (string)
function getValue(itemLink, key)
    for index, value in next, lootTable do
        if value["itemid"] == itemLink:match("item:(%d+):") then
            return value[key]
        end
    end
end

-- Handles OnTooltipSetItem
-- param: tooltip
-- More info: https://wowwiki-archive.fandom.com/wiki/UIHANDLER_OnTooltipSetItem
function injectTooltip(tooltip)
    -- https://wowwiki-archive.fandom.com/wiki/API_GameTooltip_GetItem
    -- Formatted item link (e.g. "|cff9d9d9d|Hitem:7073:0:0:0:0:0:0:0|h[Broken Fang]|h|r").
    local itemName, itemLink = tooltip:GetItem()
    
    if itemLink then
        sections = getValue(itemLink, "sections")
        --prio = getValue(itemLink, "prio")
        --note = getValue(itemLink, "note")

        if sections then
            -- We have at least one line to add
            tooltip:AddLine(string.format("\n|c000af5a2Halcyon Loot Prio       "))



            for index, value in next, sections do
                tooltip:AddLine(string.format("|c00468fcf %s [%s]", value["prio"], value["sheet"]))
            
                if value["note"] then
                    -- We have an optional note
                    tooltip:AddLine(string.format("|c00468fcf     Note: %s", value["note"]))
                end
            end



        end
    end
end


-- Add hooks
GameTooltip:HookScript("OnTooltipSetItem", injectTooltip)
ItemRefTooltip:HookScript("OnTooltipSetItem", injectTooltip)
ShoppingTooltip:HookScript("OnTooltipSetItem", injectTooltip)


