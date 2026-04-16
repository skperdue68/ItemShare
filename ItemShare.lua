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

local function GetCurrentAccountName()
    return GetDisplayName() or "@Unknown"
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

local function BuildItemKey(itemName, itemTypeName, quality, traitName)
    return string.format("%s||%s||%d||%s", itemName or "", itemTypeName or "", quality or 0, traitName or "")
end

local function GetItemShareKeyFromBagSlot(bagId, slotIndex)
    local itemName = GetNormalizedItemName(bagId, slotIndex)
    if not itemName then
        return nil, nil, nil, nil, nil, nil
    end

    local itemType, itemTypeName = GetItemTypeInfo(bagId, slotIndex)
    local quality, qualityName = GetItemQualityInfo(bagId, slotIndex)
    local traitName, traitType = GetItemTraitInfo(bagId, slotIndex)
    local itemKey = BuildItemKey(itemName, itemTypeName, quality, traitName)

    return itemKey, itemName, itemType, itemTypeName, quality, qualityName, traitName, traitType
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

local function AddItemToShareFromBagSlot(bagId, slotIndex)
    local isEligible, reason = IsItemEligibleForShare(bagId, slotIndex)
    if not isEligible then
        dmsg(string.format("Cannot add item to share: %s", tostring(reason)))
        return
    end

    local itemKey, itemName, itemType, itemTypeName, quality, qualityName, traitName, traitType =
        GetItemShareKeyFromBagSlot(bagId, slotIndex)

    if not itemKey or not itemName then
        dmsg("Could not determine item name.")
        return
    end

    dmsg(string.format("Add to Share triggered for item: %s", itemName))

    local accountName = GetCurrentAccountName()
    local incrementCount = GetItemShareCount(bagId, slotIndex)
    local now = GetTimeStamp()

    local existing = ItemShare.savedVars.sharedItems[itemKey]
    if existing then
        existing.accountName = accountName
        existing.itemName = itemName
        existing.itemType = itemType
        existing.itemTypeName = itemTypeName
        existing.quality = quality
        existing.qualityName = qualityName
        existing.trait = traitName
        existing.traitType = traitType
        existing.count = (existing.count or 0) + incrementCount
        existing.lastDumpedAt = now
        dmsg(string.format(
            "Updated %s [%s, %s%s] (count: %d)",
            itemName,
            itemTypeName,
            qualityName,
            traitName ~= "" and (", " .. traitName) or "",
            existing.count
        ))
    else
        ItemShare.savedVars.sharedItems[itemKey] = {
            itemName = itemName,
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
            "Added %s [%s, %s%s] to share (count: %d)",
            itemName,
            itemTypeName,
            qualityName,
            traitName ~= "" and (", " .. traitName) or "",
            incrementCount
        ))
    end
end

local function RemoveItemFromShareFromBagSlot(bagId, slotIndex)
    local itemKey, itemName, _, itemTypeName, _, qualityName, traitName =
        GetItemShareKeyFromBagSlot(bagId, slotIndex)

    if not itemKey or not itemName then
        dmsg("Could not determine item name.")
        return
    end

    dmsg(string.format("Remove from Share triggered for item: %s", itemName))

    if ItemShare.savedVars.sharedItems[itemKey] then
        ItemShare.savedVars.sharedItems[itemKey] = nil
        dmsg(string.format(
            "Removed %s [%s, %s%s] from shared list.",
            itemName,
            tostring(itemTypeName or ""),
            tostring(qualityName or ""),
            traitName ~= "" and (", " .. traitName) or ""
        ))
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

    local itemKey, itemName = GetItemShareKeyFromBagSlot(bagId, slotIndex)
    if not itemKey or not itemName then
        return
    end

    local isInSharedList = ItemShare.savedVars.sharedItems[itemKey] ~= nil
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
            "%s | type=%s | quality=%s | trait=%s | count=%d | account=%s",
            tostring(entry.itemName or ""),
            tostring(entry.itemTypeName or entry.itemType or ""),
            tostring(entry.qualityName or entry.quality or ""),
            (entry.trait and entry.trait ~= "") and entry.trait or "-",
            tonumber(entry.count or 0),
            tostring(entry.accountName or "@Unknown")
        ))
    end
    dmsg(string.format("Total unique shared items: %d", total))
end

local function OnSlashCommand(arg)
    arg = zo_strlower(arg or "")

    if arg == "list" then
        PrintSharedItems()
    elseif arg == "clear" or arg == "reset" then
        ResetSharedItems()
    else
        dmsg("Commands:")
        dmsg("/itemshare list  - List shared entries")
        dmsg("/itemshare reset - Clear all shared entries")
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
