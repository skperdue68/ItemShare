
ItemShare = {}
ItemShare.name = "ItemShare"
ItemShare.savedVars = nil

local defaults = {
    sharedItems = {}
}

local function dmsg(text)
    d(string.format("[ItemShare] %s", tostring(text)))
end

local function GetNormalizedItemName(bagId, slotIndex)
    local itemName = GetItemName(bagId, slotIndex)
    if not itemName or itemName == "" then
        return nil
    end
    return zo_strformat("<<t:1>>", itemName)
end

local function GetNormalizedItemLink(bagId, slotIndex)
    local itemLink = GetItemLink(bagId, slotIndex, LINK_STYLE_DEFAULT)
    if not itemLink or itemLink == "" then
        return nil
    end
    return itemLink
end

local function GetCurrentAccountName()
    return GetDisplayName() or "@Unknown"
end

local function GetBagLocationLabel(bagId)
    if bagId == BAG_BACKPACK then
        return "Inventory"
    elseif bagId == BAG_BANK then
        return "Bank"
    elseif BAG_SUBSCRIBER_BANK and bagId == BAG_SUBSCRIBER_BANK then
        return "Subscriber Bank"
    elseif BAG_VIRTUAL and bagId == BAG_VIRTUAL then
        return "Craft Bag"
    elseif BAG_HOUSE_BANK_ONE and bagId == BAG_HOUSE_BANK_ONE then
        return "House Chest 1"
    elseif BAG_HOUSE_BANK_TWO and bagId == BAG_HOUSE_BANK_TWO then
        return "House Chest 2"
    elseif BAG_HOUSE_BANK_THREE and bagId == BAG_HOUSE_BANK_THREE then
        return "House Chest 3"
    elseif BAG_HOUSE_BANK_FOUR and bagId == BAG_HOUSE_BANK_FOUR then
        return "House Chest 4"
    elseif BAG_HOUSE_BANK_FIVE and bagId == BAG_HOUSE_BANK_FIVE then
        return "House Chest 5"
    elseif BAG_HOUSE_BANK_SIX and bagId == BAG_HOUSE_BANK_SIX then
        return "House Chest 6"
    elseif BAG_HOUSE_BANK_SEVEN and bagId == BAG_HOUSE_BANK_SEVEN then
        return "House Chest 7"
    elseif BAG_HOUSE_BANK_EIGHT and bagId == BAG_HOUSE_BANK_EIGHT then
        return "House Chest 8"
    elseif BAG_HOUSE_BANK_NINE and bagId == BAG_HOUSE_BANK_NINE then
        return "House Chest 9"
    elseif BAG_HOUSE_BANK_TEN and bagId == BAG_HOUSE_BANK_TEN then
        return "House Chest 10"
    end

    return string.format("Bag %s", tostring(bagId))
end

local function JoinSortedLocations(locationSet)
    local locations = {}
    for locationName, hasLocation in pairs(locationSet or {}) do
        if hasLocation then
            table.insert(locations, locationName)
        end
    end

    table.sort(locations)
    return table.concat(locations, ", ")
end

local function GetLocalizedStringByPrefix(prefix, value, fallback)
    local globalName = prefix .. tostring(value)
    local stringId = _G[globalName]
    if stringId then
        local text = GetString(stringId)
        if text and text ~= "" then
            return zo_strformat("<<t:1>>", text)
        end
    end
    return fallback or tostring(value)
end

local function GetItemTypeInfo(bagId, slotIndex)
    local itemType = GetItemType(bagId, slotIndex) or 0
    local itemTypeName = GetLocalizedStringByPrefix("SI_ITEMTYPE", itemType, tostring(itemType))
    return itemType, itemTypeName
end

local function GetItemQualityInfo(bagId, slotIndex)
    local quality = GetItemQuality(bagId, slotIndex) or 0
    local qualityName = GetLocalizedStringByPrefix("SI_ITEMQUALITY", quality, tostring(quality))
    return quality, qualityName
end

local function GetItemTraitInfo(bagId, slotIndex)
    local traitType = GetItemTrait(bagId, slotIndex)

    if not traitType or traitType == ITEM_TRAIT_TYPE_NONE or traitType == 0 then
        return "", ITEM_TRAIT_TYPE_NONE
    end

    local traitName = GetLocalizedStringByPrefix("SI_ITEMTRAITTYPE", traitType, tostring(traitType))
    return traitName, traitType
end

local function GetItemShareCount(bagId, slotIndex)
    local stackCount = GetSlotStackSize(bagId, slotIndex)
    if stackCount and stackCount > 0 then
        return stackCount
    end
    return 1
end

local function BuildItemKey(itemName, itemTypeName, quality, traitName, itemLink)
    return string.format(
        "%s||%s||%d||%s||%s",
        itemName or "",
        itemTypeName or "",
        quality or 0,
        traitName or "",
        itemLink or ""
    )
end

local function BuildLegacyItemKey(itemName, itemTypeName, quality, traitName)
    return string.format("%s||%s||%d||%s", itemName or "", itemTypeName or "", quality or 0, traitName or "")
end

local function GetItemShareKeyFromBagSlot(bagId, slotIndex)
    local itemName = GetNormalizedItemName(bagId, slotIndex)
    if not itemName then
        return nil, nil, nil, nil, nil, nil
    end

    local itemLink = GetNormalizedItemLink(bagId, slotIndex)
    local itemType, itemTypeName = GetItemTypeInfo(bagId, slotIndex)
    local quality, qualityName = GetItemQualityInfo(bagId, slotIndex)
    local traitName, traitType = GetItemTraitInfo(bagId, slotIndex)
    local itemKey = BuildItemKey(itemName, itemTypeName, quality, traitName, itemLink)
    local legacyItemKey = BuildLegacyItemKey(itemName, itemTypeName, quality, traitName)

    return itemKey, itemName, itemType, itemTypeName, quality, qualityName, traitName, traitType, itemLink, legacyItemKey
end

local function IsItemEligibleForShare(bagId, slotIndex)
    if bagId == nil or slotIndex == nil then
        return false, "Invalid inventory slot."
    end

    if IsItemBound(bagId, slotIndex) then
        return false, "Item is already bound."
    end

    if IsItemBoPAndTradeable(bagId, slotIndex) then
        return false, "Item is restricted to temporary group sharing."
    end

    return true, nil
end

local function FindExistingSharedEntry(itemKey, legacyItemKey)
    local entry = ItemShare.savedVars.sharedItems[itemKey]
    if entry then
        return entry, itemKey
    end

    if legacyItemKey and legacyItemKey ~= itemKey then
        entry = ItemShare.savedVars.sharedItems[legacyItemKey]
        if entry then
            return entry, legacyItemKey
        end
    end

    return nil, nil
end

local function MigrateEntryKeyIfNeeded(oldKey, newKey, entry)
    if oldKey and newKey and entry and oldKey ~= newKey then
        ItemShare.savedVars.sharedItems[oldKey] = nil
        ItemShare.savedVars.sharedItems[newKey] = entry
    end
end

local function AddItemToShareFromBagSlot(bagId, slotIndex)
    local isEligible, reason = IsItemEligibleForShare(bagId, slotIndex)
    if not isEligible then
        dmsg(string.format("Cannot add item to share: %s", tostring(reason)))
        return
    end

    local itemKey, itemName, itemType, itemTypeName, quality, qualityName, traitName, traitType, itemLink, legacyItemKey =
        GetItemShareKeyFromBagSlot(bagId, slotIndex)

    if not itemKey or not itemName then
        dmsg("Could not determine item name.")
        return
    end

    dmsg(string.format("Add to Share triggered for item: %s", itemName))

    local accountName = GetCurrentAccountName()
    local incrementCount = GetItemShareCount(bagId, slotIndex)
    local sharedFrom = GetBagLocationLabel(bagId)
    local now = GetTimeStamp()

    local existing, existingKey = FindExistingSharedEntry(itemKey, legacyItemKey)
    if existing then
        existing.accountName = accountName
        existing.itemName = itemName
        existing.itemType = itemType
        existing.itemTypeName = itemTypeName
        existing.quality = quality
        existing.qualityName = qualityName
        existing.trait = traitName
        existing.traitType = traitType
        existing.itemLink = itemLink
        existing.count = (existing.count or 0) + incrementCount
        existing.lastDumpedAt = now

        MigrateEntryKeyIfNeeded(existingKey, itemKey, existing)

        dmsg(string.format(
            "Updated %s [%s, %s%s] from %s (count: %d)",
            itemName,
            itemTypeName,
            qualityName,
            traitName ~= "" and (", " .. traitName) or "",
            sharedFrom,
            existing.count
        ))
    else
        ItemShare.savedVars.sharedItems[itemKey] = {
            itemName = itemName,
            itemLink = itemLink,
            accountName = accountName,
            itemType = itemType,
            itemTypeName = itemTypeName,
            quality = quality,
            qualityName = qualityName,
            trait = traitName,
            traitType = traitType,
            count = incrementCount,
            firstDumpedAt = now,
            lastDumpedAt = now,
        }
        dmsg(string.format(
            "Added %s [%s, %s%s] to share from %s (count: %d)",
            itemName,
            itemTypeName,
            qualityName,
            traitName ~= "" and (", " .. traitName) or "",
            sharedFrom,
            incrementCount
        ))
    end
end

local function RemoveItemFromShareByKey(itemKey, entry)
    if not itemKey or not entry then
        return false
    end

    ItemShare.savedVars.sharedItems[itemKey] = nil
    dmsg(string.format(
        "Removed %s [%s, %s%s] from shared list.",
        tostring(entry.itemName or ""),
        tostring(entry.itemTypeName or entry.itemType or ""),
        tostring(entry.qualityName or entry.quality or ""),
        (entry.trait and entry.trait ~= "") and (", " .. entry.trait) or ""
    ))
    return true
end

local function RemoveItemFromShareFromBagSlot(bagId, slotIndex)
    local itemKey, itemName, _, itemTypeName, _, qualityName, traitName, _, _, legacyItemKey =
        GetItemShareKeyFromBagSlot(bagId, slotIndex)

    if not itemKey or not itemName then
        dmsg("Could not determine item name.")
        return
    end

    dmsg(string.format("Remove from Share triggered for item: %s", itemName))

    local existing, existingKey = FindExistingSharedEntry(itemKey, legacyItemKey)
    if existing then
        RemoveItemFromShareByKey(existingKey, existing)
    else
        dmsg(string.format("Item was not found in shared list: %s", itemName))
    end
end

local function AddInventoryContextMenu(slot)
    if not slot or slot.bagId == nil or slot.slotIndex == nil then
        return
    end

    local bagId = slot.bagId
    local slotIndex = slot.slotIndex

    local itemKey, itemName, _, _, _, _, _, _, _, legacyItemKey = GetItemShareKeyFromBagSlot(bagId, slotIndex)
    if not itemKey or not itemName then
        return
    end

    local existing = FindExistingSharedEntry(itemKey, legacyItemKey)
    local isInSharedList = existing ~= nil
    if isInSharedList then
        AddCustomMenuItem("Remove from Shared List", function()
            RemoveItemFromShareFromBagSlot(bagId, slotIndex)
        end, MENU_ADD_OPTION_LABEL)
        return
    end

    local isEligible = IsItemEligibleForShare(bagId, slotIndex)
    if not isEligible then
        return
    end

    AddCustomMenuItem("Add to Share", function()
        AddItemToShareFromBagSlot(bagId, slotIndex)
    end, MENU_ADD_OPTION_LABEL)
end

local function RegisterContextMenus()
    if not LibCustomMenu then
        dmsg("LibCustomMenu not found.")
        return
    end

    LibCustomMenu:RegisterContextMenu(function(inventorySlot)
        AddInventoryContextMenu(inventorySlot)
    end, LibCustomMenu.CATEGORY_LATE)
end

local function ResetSharedItems()
    ItemShare.savedVars.sharedItems = {}
    dmsg("Saved variables reset. No shared items remain.")
end

local function PrintSharedItems()
    local total = 0
    for _, entry in pairs(ItemShare.savedVars.sharedItems) do
        total = total + 1
        dmsg(string.format(
            "%s | type=%s | quality=%s | trait=%s | count=%d | account=%s%s",
            tostring(entry.itemName or ""),
            tostring(entry.itemTypeName or entry.itemType or ""),
            tostring(entry.qualityName or entry.quality or ""),
            (entry.trait and entry.trait ~= "") and entry.trait or "-",
            tonumber(entry.count or 0),
            tostring(entry.accountName or "@Unknown"),
            entry.itemLink and entry.itemLink ~= "" and (" | link=" .. entry.itemLink) or ""
        ))
    end
    dmsg(string.format("Total unique shared items: %d", total))
end

local function GetTrackedBagIds()
    local bagIds = {
        BAG_BACKPACK,
        BAG_BANK,
    }

    if BAG_SUBSCRIBER_BANK then
        table.insert(bagIds, BAG_SUBSCRIBER_BANK)
    end

    if BAG_VIRTUAL then
        table.insert(bagIds, BAG_VIRTUAL)
    end

    local houseBanks = {
        BAG_HOUSE_BANK_ONE,
        BAG_HOUSE_BANK_TWO,
        BAG_HOUSE_BANK_THREE,
        BAG_HOUSE_BANK_FOUR,
        BAG_HOUSE_BANK_FIVE,
        BAG_HOUSE_BANK_SIX,
        BAG_HOUSE_BANK_SEVEN,
        BAG_HOUSE_BANK_EIGHT,
        BAG_HOUSE_BANK_NINE,
        BAG_HOUSE_BANK_TEN,
    }

    for _, bagId in ipairs(houseBanks) do
        if bagId then
            table.insert(bagIds, bagId)
        end
    end

    return bagIds
end

local function BuildActualInventoryIndex()
    local countsByLink = {}
    local countsByLegacyKey = {}
    local uniqueLinkByLegacyKey = {}
    local locationsByLink = {}
    local locationsByLegacyKey = {}

    for _, bagId in ipairs(GetTrackedBagIds()) do
        local bagLocation = GetBagLocationLabel(bagId)
        local bagSize = GetBagSize(bagId) or 0
        for slotIndex = 0, bagSize - 1 do
            if not IsSlotEmpty(bagId, slotIndex) then
                local isEligible = IsItemEligibleForShare(bagId, slotIndex)
                if isEligible then
                    local itemKey, itemName, _, itemTypeName, quality, _, traitName, _, itemLink, legacyItemKey =
                        GetItemShareKeyFromBagSlot(bagId, slotIndex)

                    local stackCount = GetItemShareCount(bagId, slotIndex)

                    if itemLink and itemLink ~= "" then
                        countsByLink[itemLink] = (countsByLink[itemLink] or 0) + stackCount
                        locationsByLink[itemLink] = locationsByLink[itemLink] or {}
                        locationsByLink[itemLink][bagLocation] = true
                    end

                    local effectiveLegacyKey = legacyItemKey
                    if (not effectiveLegacyKey or effectiveLegacyKey == "") and itemName then
                        effectiveLegacyKey = BuildLegacyItemKey(itemName, itemTypeName, quality, traitName)
                    end

                    if effectiveLegacyKey and effectiveLegacyKey ~= "" then
                        countsByLegacyKey[effectiveLegacyKey] = (countsByLegacyKey[effectiveLegacyKey] or 0) + stackCount
                        locationsByLegacyKey[effectiveLegacyKey] = locationsByLegacyKey[effectiveLegacyKey] or {}
                        locationsByLegacyKey[effectiveLegacyKey][bagLocation] = true

                        if itemLink and itemLink ~= "" then
                            local currentUniqueLink = uniqueLinkByLegacyKey[effectiveLegacyKey]
                            if currentUniqueLink == nil then
                                uniqueLinkByLegacyKey[effectiveLegacyKey] = itemLink
                            elseif currentUniqueLink ~= itemLink then
                                uniqueLinkByLegacyKey[effectiveLegacyKey] = false
                            end
                        end
                    end
                end
            end
        end
    end

    return countsByLink, countsByLegacyKey, uniqueLinkByLegacyKey, locationsByLink, locationsByLegacyKey
end

local function SyncSharedItemsWithInventory()
    local sharedItems = ItemShare.savedVars.sharedItems
    local countsByLink, countsByLegacyKey, uniqueLinkByLegacyKey, locationsByLink, locationsByLegacyKey = BuildActualInventoryIndex()
    local removedCount = 0
    local updatedCount = 0
    local unchangedCount = 0
    local migratedCount = 0

    local snapshot = {}
    for itemKey, entry in pairs(sharedItems) do
        snapshot[itemKey] = entry
    end

    for itemKey, entry in pairs(snapshot) do
        local itemName = tostring(entry.itemName or "Unknown Item")
        local actualCount = 0
        local currentLocation = ""
        local matchSource = nil

        if entry.itemLink and entry.itemLink ~= "" then
            actualCount = countsByLink[entry.itemLink] or 0
            currentLocation = JoinSortedLocations(locationsByLink[entry.itemLink])
            matchSource = "itemLink"
        else
            local legacyKey = BuildLegacyItemKey(
                entry.itemName,
                entry.itemTypeName,
                entry.quality,
                entry.trait
            )
            actualCount = countsByLegacyKey[legacyKey] or 0
            currentLocation = JoinSortedLocations(locationsByLegacyKey[legacyKey])
            matchSource = "legacyKey"

            local discoveredLink = uniqueLinkByLegacyKey[legacyKey]
            if type(discoveredLink) == "string" and discoveredLink ~= "" then
                entry.itemLink = discoveredLink
                currentLocation = JoinSortedLocations(locationsByLink[discoveredLink]) or currentLocation
                matchSource = "legacyKey+itemLink"
            end
        end

        if actualCount <= 0 then
            RemoveItemFromShareByKey(itemKey, entry)
            removedCount = removedCount + 1
            dmsg(string.format("Sync removed %s because it was not found in inventory, bank, or house storage.", itemName))
        else
            local desiredKey = BuildItemKey(
                entry.itemName,
                entry.itemTypeName,
                entry.quality,
                entry.trait,
                entry.itemLink
            )

            if desiredKey ~= itemKey then
                sharedItems[itemKey] = nil
                sharedItems[desiredKey] = entry
                itemKey = desiredKey
                migratedCount = migratedCount + 1
            end

            local hadChanges = false

            if tonumber(entry.count or 0) ~= actualCount then
                entry.count = actualCount
                hadChanges = true
            end

            if tostring(entry.sharedFrom or "") ~= tostring(currentLocation or "") then
                entry.sharedFrom = currentLocation
                hadChanges = true
            end

            if hadChanges then
                entry.lastDumpedAt = GetTimeStamp()
                updatedCount = updatedCount + 1
                dmsg(string.format(
                    "Sync updated %s count to %d and location to %s using %s match.",
                    itemName,
                    actualCount,
                    tostring(currentLocation ~= "" and currentLocation or "-"),
                    tostring(matchSource)
                ))
            else
                unchangedCount = unchangedCount + 1
            end
        end
    end

    dmsg(string.format(
        "Sync complete. Removed=%d Updated=%d Unchanged=%d MigratedKeys=%d",
        removedCount,
        updatedCount,
        unchangedCount,
        migratedCount
    ))
end

local function OnSlashCommand(arg)
    arg = zo_strlower(arg or "")

    if arg == "list" then
        PrintSharedItems()
    elseif arg == "clear" or arg == "reset" then
        ResetSharedItems()
    elseif arg == "sync" or arg == "cleanup" or arg == "reconcile" then
        SyncSharedItemsWithInventory()
    else
        dmsg("Commands:")
        dmsg("/itemshare list    - List shared entries")
        dmsg("/itemshare reset   - Clear all shared entries")
        dmsg("/itemshare sync    - Update counts and remove items no longer owned")
    end
end

local function OnAddonLoaded(event, addonName)
    if addonName ~= ItemShare.name then
        return
    end

    EVENT_MANAGER:UnregisterForEvent(ItemShare.name, EVENT_ADD_ON_LOADED)

    ItemShare.savedVars = ZO_SavedVars:NewAccountWide(
        "ItemShareSavedVars",
        1,
        nil,
        defaults
    )

    RegisterContextMenus()

    SLASH_COMMANDS["/itemshare"] = OnSlashCommand
    SLASH_COMMANDS["/ishare"] = OnSlashCommand

    dmsg("Loaded.")
end

EVENT_MANAGER:RegisterForEvent(ItemShare.name, EVENT_ADD_ON_LOADED, OnAddonLoaded)
