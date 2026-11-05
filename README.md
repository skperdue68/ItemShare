Requires LibCustomMenu

Download and extract to your ESO AddOns Directory
Ensure Add on is loaded at startup


You can use commands 

/ishare or /itemshare to control functionality

/ishare list - shows you the list of items that are set to being shared.  You may remove or re-add items that are in this list from this panel
/ishare reset - Clears all items from your shared list
/ishare save - Does a reloadUI which forces the Shared list to be written to disk so you can import it to the spreadsheet
/ishare sync - Checks your shared list against items you still have, items you no longer have are removed from the shared list. Always-shared items are left alone, because a standing offer is meant to outlive the copy.

You may add items to the shared list from Inventory, House Storage, Your Bank, or the Furnishings Vault, just right click the item and select "Add to Share" or "Remove from Share" as you desire.

Items that are bound to you cannot be shared, and neither can items the game only lets you hand to your group for a short while (bind on pickup, still tradeable). The addon refuses both with a message instead of adding them, and it checks again when an item is re-added from the share list window, because an item can become bound after it was filed.

Right-click an item and choose "Always Share This Item" to make it a standing offer: it stays offered even when you no longer have a copy, and `/ishare sync` never removes it. A normally shared item is different - it is only offered while it is in a bag the addon can read, and it disappears from the list once it is gone.

The badge on an item icon shows which of the two it is. A green gift means it is shared right now; the same gift in gold with a closed ring means it is always shared, and that gold badge stays on the item even when you have none in your bags. Hovering the badge says which state it is in.

Once Everything is as you want it, save the file by doing /ishare save or typing /reloadui

Open your google spreadsheet (Documentation includes a link to a templatae file you can use which should have everything set for you, just save it as a copy yourself), and use the update feature to add your data into the spreadsheet, its that easy.
