# ItemShare

An ESO addon that marks tradeable items as shareable and writes them to `ItemShareSavedVars`, which the ItemShare website reads. This repository holds the addon only; the website is a separate project.

Requires LibCustomMenu

Download and extract to your ESO AddOns Directory
Ensure Add on is loaded at startup


You can use commands 

/ishare or /itemshare to control functionality

/ishare list - opens or closes the shared items window. Right-click a row there to remove or re-add a share.
/ishare reset (or clear) - clears every share and every standing offer.
/ishare sync (or cleanup, reconcile) - recounts your shares against the bags that are readable now and drops the ones you no longer have. Standing offers are left alone, because an offer is meant to outlive the copy.
/ishare rebuild - rewrites every share from your current bags; it also runs on login and when a bank opens or closes.
/ishare save - rebuilds and then offers to reload the UI, which is what writes the file to disk.
/ishare debug - turns the addon's own status messages on or off.
/ishare guilds - lists the guilds captured for the current character.
/ishare always - lists the items you always offer, whether or not a copy is in your bags.
/ishare noalways <part of the name> - stops always sharing one item.
/ishare reserve <amount> <part of the name> - keeps that many back and offers the rest.
/ishare noreserve <part of the name> - shares the whole stack again.
/ishare reserves - lists every item with a reserve, and what it is offering.
/ishare guildbank - toggles sharing from the guild bank.

The window and the item menus cover the same ground without typing:

- Right-click an item in a bag, the bank, house storage or the furnishing vault: **Add to Share**, **Remove from Shared List**, **Always Share This Item**, **Stop Always Sharing**, **Share Reserve** (asks for a number of at least 1) and, when a reserve is set, **Share Reserve: Keep N Back...** and **Clear Share Reserve**.
- In `/ishare list`, right-click a row to stop sharing it, or to stop the standing offer a row stands for.

You may add items to the shared list from Inventory, House Storage, Your Bank, or the Furnishings Vault, just right click the item and select "Add to Share" or "Remove from Share" as you desire.

Items that are bound to you cannot be shared, and neither can items the game only lets you hand to your group for a short while (bind on pickup, still tradeable). The addon refuses both with a message instead of adding them, and it checks again when an item is re-added from the share list window, because an item can become bound after it was filed.

Right-click an item and choose "Always Share This Item" to make it a standing offer: it stays offered even when you no longer have a copy, and `/ishare sync` never removes it. A normally shared item is different - it is only offered while it is in a bag the addon can read, and it disappears from the list once it is gone.

The badge on an item icon shows which of the two it is. A green gift means it is shared right now; the same gift in gold with a closed ring means it is always shared, and that gold badge stays on the item even when you have none in your bags. A small amber padlock tag is added to either badge when the item has a share reserve. An item that has a reserve but is not shared at all shows the padlock on its own, so a reserve left behind on something you stopped sharing is still visible - and can be cleared - instead of being forgotten. Hovering the badge says which state it is in and, when there is a reserve, how much is being kept back.

`/ishare list` shows what is on offer right now: every share with something above its reserve, plus any standing offer that has a copy in your bags even if that copy was never marked by hand. A row that offers nothing is not listed at all, a row with a reserve says `keep N` beside its location, and right-clicking a row that only stands for a standing offer stops the offer.

### Share reserve: share only what you have above a number

If you want to keep some of a stack for yourself, give it a share reserve. The addon then offers only the quantity above that number, and the share list recalculates whenever your inventory changes:

```
/itemshare reserve 12 soul gem     keep 12, so 20 held shares 8
/itemshare reserve 12 soul gem     used 15 later, so 3 are shared
/itemshare reserve 12 soul gem     used down to 10 later, so nothing is shared
/itemshare reserves               list everything that has a reserve
/itemshare noreserve soul gem     share the whole stack again
```

The same reserve can be set from the item's right-click menu: **Share Reserve** opens a number box with the item already chosen, refuses anything that is not a whole number of at least 1, and offers **Clear Reserve** once something is kept back.

The name only has to be part of the item name, and it is refused if it matches more than one shared item. A share that is entirely inside its reserve stays marked at zero, so it comes back on its own when you acquire more - nothing has to be re-marked. Items with no reserve behave exactly as before. A reserve covers standing offers too: the list only shows one while you hold more than the reserve, and the offer comes back by itself once you do.

Once Everything is as you want it, save the file by doing /ishare save or typing /reloadui

Then open the ItemShare site and upload the file there. The item list, the requests and everything that becomes of them happen on the site.
