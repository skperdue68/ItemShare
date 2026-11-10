# ESOUI addon page

The addon's page on ESOUI carries two pieces of text: a short description and a full
description written in ESOUI's BBCode. Both are kept here so the page can be rewritten
when the addon changes, instead of drifting away from it. The full description
describes behaviour, so a change to a command, a badge, a location or the website
belongs here too.

The site the addon feeds is https://itemshare.perdues.me; its source is
https://github.com/skperdue68/ItemShare-Web.

## Short description

Mark items in your bags, bank, house storage or furnishing vault as shareable, then
upload the file the addon writes to itemshare.perdues.me so guild members can see what
you are offering and request it.

## Full description

Uses `[b]`, `[i]`, `[size]`, `[color=#hex]`, `[url]`, `[list]` with `[*]`, `[code]` and
`[hr]` only, all of which ESOUI renders. Paste the block as it stands into the
description field.

```bbcode
[b][size=5][color=#D4AF37]ItemShare[/color][/size][/b]
[color=#9AA0A6][i]Track what you are willing to give away. Let the website handle who asks for it.[/i][/color]

[hr]

ItemShare keeps track of the tradeable items you are willing to hand out, so you do not have to leave things in a guild bank, post lists in chat, or remember what you promised whom. You mark an item with one right-click, the addon records it in a saved file, and you upload that file to [url=https://itemshare.perdues.me]itemshare.perdues.me[/url], where other players browse what is on offer and request what they need.

[color=#E3B341][b]Requires LibCustomMenu.[/b][/color] Please install it first. The right-click entries that do all the work come from that library, and without it the addon has very little it can do.

[hr]

[b][size=4][color=#D4AF37]How it works[/color][/size][/b]

[list]
[*]Right-click an item in your bags, your bank, a house storage chest or the furnishing vault and choose [b]Add to Share[/b].
[*]The item goes into [b]ItemShareSavedVars[/b], an account-wide saved-variables file, along with its link, name, quality, trait, type, stack count, where it is, the character it came from and that character's guilds.
[*]The addon keeps that file current by itself. It rewrites your shares a few seconds after a character loads and whenever a bank opens or closes, and it drops shares it can prove are gone from a bag it is able to read.
[*]When the list is the way you want it, [b]/ishare save[/b] (or /reloadui) writes the file to disk.
[*]Upload the file at [url=https://itemshare.perdues.me]itemshare.perdues.me[/url]. Browsing, requests and everything that follows a request happen on the site. The addon does not handle requests in game.
[/list]

[b][size=4][color=#D4AF37]What can be shared, and from where[/color][/size][/b]

[list]
[*][b]Inventory[/b] - recorded per character, so each character keeps its own rows.
[*][b]Bank[/b] and [b]Craft Bag[/b].
[*][b]House storage[/b] - recorded under the name of the chest, not its number.
[*][b]Furnishing vault[/b].
[*][b]Guild bank[/b] - off by default. /ishare guildbank toggles it.
[/list]

[b]House storage by name.[/b] ESO lets you rename each of the ten storage chests, so a row reads [i]Materials[/i] instead of [i]House Storage 8[/i]. The chest's own name is used when you have not renamed it, and the slot number only when neither is available. Rename a chest and the next rebuild relabels its shares, and two chests given the same name stay separate rows.

[b]Furnishings for sharing live in your storage house.[/b] Stand in that house and run [b]/ishare scan[/b]: every furnishing placed in it is collected, and each piece that is not bound to you is offered as [i]Placed (Mournoth Keep)[/i], one row per furnishing with the number of copies standing there. Running it again is the record for that house - a piece you picked up is taken back, and one carried in is added. Only the house you are standing in can be read, so no other house is ever changed, and furnishings in the vault or in a bag are shared the usual way. Bound furnishings are counted and skipped.

[b]Bound items are refused[/b], and so are items the game only lets you hand to your group for a short while (bind on pickup, still tradeable). The addon says so instead of adding them, and checks again when an item is re-added from the shared window, because an item can become bound after it was filed.

[b][size=4][color=#D4AF37]The shared items window[/color][/size][/b]

[b]/ishare list[/b] opens a movable, resizable window that remembers its position and size.

[list]
[*]One row per item per place. Stacks of the same thing in one bag are added together, so two stacks of 200 soul gems in the bank are 400 on offer, but a share in your bags and a share in the bank stay two rows, each naming its own place and counting only what is there.
[*]Left-click an item name for its full tooltip.
[*]Right-click a row to step it on: not shared, shared, always shared, and back to not shared.
[*]A row that offers nothing is not listed at all, and a row with a reserve says [i]keep N[/i] beside its location.
[*]The window reads the counts saved with your shares rather than measuring the bags again, which is why the bank's row is still there while the bank is shut.
[/list]

[b][size=4][color=#D4AF37]Standing offers[/color][/size][/b]

[color=#E3B341][b]Always Share This Item[/b][/color] on an item's menu makes it a standing offer. It stays offered even when you hold no copy, and [b]/ishare sync[/b] never removes it. A normally shared item is different: it is offered only while it is in a bag the addon can read, and it disappears from the list once it is gone.

[b]/ishare always[/b] lists your standing offers, and [b]/ishare noalways <part of the name>[/b] ends one.

[b][size=4][color=#D4AF37]Share reserve: offer only what you have above a number[/color][/size][/b]

Keep part of a stack for yourself and offer the rest:

[code]/ishare reserve 12 soul gem    keep 12, so holding 20 shares 8
/ishare reserve 12 soul gem    used 15 later, so 3 are shared
/ishare reserve 12 soul gem    used down to 10 later, so nothing is shared
/ishare noreserve soul gem     share the whole stack again[/code]

The same thing is on the item's right-click menu as [b]Share Reserve[/b], which asks for a number of at least 1. A reserve belongs to the item in one place, so the stack in your bank and the stack in your bags keep their own numbers. A share entirely inside its reserve stays marked at zero and comes back by itself when you acquire more, so nothing has to be re-marked. The name only has to be part of the item name, and is refused if it matches more than one shared item.

[b][size=4][color=#D4AF37]Badges[/color][/size][/b]

[list]
[*]A [color=#5BA85B][b]green gift[/b][/color] on an item icon means it is shared right now.
[*]The same gift in [color=#D4AF37][b]gold[/b][/color], with a closed ring, means it is always shared, and the badge stays on the item even when you hold none.
[*]A small [color=#E3B341][b]amber padlock[/b][/color] tag is added to either badge when the item has a share reserve. An item with a reserve that is not shared at all shows the padlock on its own, so a reserve left behind can be seen and cleared instead of forgotten.
[*]Hovering a badge says which state the item is in and, with a reserve, how much is being kept back.
[/list]

[b][size=4][color=#D4AF37]Guilds and server[/color][/size][/b]

[list]
[*]Shares are filed against the guild's id, and the file records all five guild slots for each character, writing 0 in every slot with no guild in it.
[*]On login the addon removes shares and reserves filed against a guild you are no longer in, since those banks can never be opened again. A guild's name is forgotten once no character on the account is still in it.
[*][b]/ishare guilds[/b] shows all five slots, empty ones included.
[*]The file records which platform and which megaserver the account plays on, and the site matches on them. A file written on the North American megaserver is only offered to readers on the same platform and megaserver, because nothing trades between them.
[/list]

[hr]

[b][size=4][color=#D4AF37]Commands[/color][/size][/b]

[b]/itemshare[/b] or [b]/ishare[/b]

[list]
[*][b]list[/b] - open or close the shared items window
[*][b]reset[/b] (or clear) - clear every share and every standing offer
[*][b]sync[/b] (cleanup, reconcile) - recount shares against the bags readable now and drop the ones you no longer have. Standing offers are left alone.
[*][b]rebuild[/b] - rewrite every share from your current bags. Also runs on login and when a bank opens or closes.
[*][b]scan[/b] - share every non-bound furnishing placed in the house you are standing in, and take back the ones no longer there
[*][b]save[/b] - rebuild, then offer to reload the UI, which writes the file to disk
[*][b]debug[/b] - turn the addon's status messages on or off
[*][b]guilds[/b] - list the guilds captured for the current character
[*][b]always[/b] - list the items you always offer
[*][b]noalways <name>[/b] - stop always sharing one item
[*][b]reserve <n> <name>[/b] - keep n back and offer the rest
[*][b]noreserve <name>[/b] - share the whole stack again
[*][b]reserves[/b] - list every item with a reserve
[*][b]guildbank[/b] - toggle sharing from the guild bank
[*][b]server[/b] - show the platform and megaserver this file is recorded for
[/list]

[hr]

[b][size=4][color=#D4AF37]The website[/color][/size][/b]

[url=https://itemshare.perdues.me]itemshare.perdues.me[/url] is open to anyone. Upload your ItemShareSavedVars.lua, then browse what other players are offering and request what you need. There is a demo mode if you want to look around before uploading your own file.

If you would rather run your own copy, the site's source is at [url=https://github.com/skperdue68/ItemShare-Web]github.com/skperdue68/ItemShare-Web[/url]. Nothing in the addon needs it. The addon only writes the saved file.

[b]Google Sheets:[/b] earlier versions of ItemShare exported to Google Sheets. That is gone and is not supported. The addon now writes ItemShareSavedVars.lua and the website does the rest.

[b][size=4][color=#D4AF37]Notes[/color][/size][/b]

[list]
[*]Version 1.11.7, API 101051. Requires LibCustomMenu 7.3.0 or newer.
[*]The saved file has a format version, and a file written by another one is not carried forward. Your shares come back from your bags on the next rebuild, but standing offers and reserves are set by hand rather than read from a bag, so they start empty.
[*]Parts of the addon were written with AI assistance and it is still being tidied up.
[*]Addon source: [url=https://github.com/skperdue68/ItemShare]github.com/skperdue68/ItemShare[/url]
[/list]
```
