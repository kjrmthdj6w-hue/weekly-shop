# Weekly Shop — setup guide

A phone-friendly web app built from your `Food_Improved_v4.xlsm` spreadsheet and its
`RefreshShoppingList` macro. It's plain HTML/JavaScript, so it hosts free on GitHub Pages
and installs on your iPhone Home Screen. No Apple developer account, nothing expires.

## The five pages

| Page | What it does |
|---|---|
| **Shopping list** | The aisle-sorted list the macro used to write into columns H–J. Tick items off as you shop. Duplicates show as ×2, ×3. Items it can't place are highlighted yellow with a dropdown to choose an aisle (remembered for next time). **Share list** sends it to a message/Notes. |
| **Meal planner** | Up to 8 meals a week, each with name, recipe book, page and ingredients (one per line). Below them, your **Extras & household** list (the old Misc column): tap an item to highlight it yellow = "needed this week". **Complete – populate shopping list** at the bottom does what the macro did. |
| **Aisle planner** | Reorder, rename, add and delete aisles (▲▼✎✕). Add keywords manually, fix unknown ingredients from the current list, search/edit all keywords. |
| **Weekly recipes** | This week's meals with book, page and ingredients. **Mark as used** strikes a meal through once you've cooked it (tap again to undo); **Remove** takes it out of the week's plan. Neither changes the shopping list. |
| **Previous meals** | Every meal you've planned is saved here automatically when you tap *Complete – populate shopping list*: name, recipe book, page and ingredients. Ingredients **accumulate and never duplicate** (Carrot one week, Rice the next → Carrot, Rice). Each meal counts once per week, and a **Top 5 recipes** list shows your most-chosen. Open a meal to edit its book/page, remove or add ingredients, add it to this week's plan, or delete it. |

Also: **Settings** (⚙︎ top right) for GitHub backup, and **Start a new week** on the Meal planner clears the meals and un-highlights the extras (your aisles and extras list stay).

## How Previous meals feeds the planner

- Type a meal name in the Meal planner and, when you finish typing, a matching previous meal's saved ingredients (and book/page if blank) fill in automatically. Delete what you don't need this week, then populate. Removing ingredients in the planner does **not** remove them from Previous meals; use the Previous meals page for that.
- While typing, tappable suggestions appear (top 5 when the name is empty). **↻ Load saved ingredients** re-fills a meal. The app never overwrites ingredients you've typed yourself without asking.
- Names match ignoring capitals and extra spaces. Existing backups stay compatible.

## Using several devices

- Open the app and tap **⟳ Update from latest backup** on the Home page to pull the newest data. It also checks automatically when you open or return to the app (only if this device has no unsaved changes, so nothing is overwritten silently).
- Backups happen straight away when you populate the shopping list, mark or remove a recipe, or change the aisle order, and about 10 seconds after other edits. Leaving the app also sends any pending change.
- The **aisle order** (and renames, keywords, extras, meals and history) are all part of the GitHub backup, so another device gets the same shop layout and the shopping list order follows it.
- In the Meal planner each ingredient shows as a chip: tap ✕ to drop it (e.g. rice you don't need this week).

## How it matches the macro

- Ingredients are lower-cased, "(have)" and "?" are removed, duplicates are counted.
- Highlighted ("yellow") extras are included; un-highlighted ones are ignored.
- Each ingredient is matched against your keywords (the Aisles sheet). Anything unmatched is flagged.
- Sorted by your aisle order, then alphabetically. Unknown items go last.
- Tested against your spreadsheet's own output: with the macro's exact rule it gave **identical results on all 35 lines**.

**One deliberate improvement:** the macro used the *first* matching keyword down the sheet, so a specific keyword you added later could never win (e.g. "bean" always beat "green beans"). The app uses the **longest matching keyword** instead. On your current data this changes two items, both towards what your own keywords intended: *stir fry veg mix* → Frozen, and *coconut milk* → Condiments.

Your sheet had 10 meal slots; you asked for a maximum of 8, so the app starts with the first 8 (Spanish saffron chicken and Cotelli weren't carried over). To change the limit, edit `MAX_MEALS = 8` at the top of `app.js`.

Not carried over: the `MoveYellowToShoppingList` macro, because tapping to highlight extras replaces it, and the empty "Breakfast" column E, which the macro read but your sheet doesn't use.

## Put it online (same as the commute tracker)

1. On github.com create a **new public repository** called `weekly-shop`.
2. **Add file → Upload files**, and upload everything in this folder: `index.html`, `app.js`, `seed.js`, `sw.js`, `manifest.json`, and the `icons` folder. Commit.
3. **Settings → Pages → Deploy from a branch → main / (root) → Save.** After a minute your address appears, like `https://YOURNAME.github.io/weekly-shop/`.
4. Open that address in **Safari** → Share → **Add to Home Screen**.

> **Don't remove and re-add the Home Screen icon later.** On iPhone each new icon gets its own empty storage, which is what wiped your commute data. Set the backup below up first.

## Turn on automatic backup (recommended)

Your data lives on the phone. This copies it to GitHub so a wipe can't lose it.

1. Create a **second repository, set to Private**, e.g. `weekly-shop-data` (so your meal plans aren't public). Add a README so it isn't empty.
2. github.com/settings/tokens → **Fine-grained tokens → Generate new token**. Limit it to **only** that private repo and set **Contents: Read and write**. Copy the token.
3. In the app: ⚙︎ **Settings** → paste the token, your GitHub username, and `weekly-shop-data` as the repo. Leave the file path as `meal-data.json` and keep "Back up automatically" on. Tap **Save**, then **Back up now**. You should see a green **✅ Backed up successfully**.
4. Don't paste the token into chat or share it. Use **Show settings text** and keep a copy in Notes or a password manager.

**How the sync behaves**
- Changes back up about 10 seconds after you stop editing, and straight away when you populate the list.
- On opening the app, if GitHub has a newer copy and this phone has no unsaved changes, it loads it automatically. After a wipe it restores itself.
- If both sides changed, it never overwrites silently. It asks, or shows a warning.
- **Load saved data** (Settings) fetches the GitHub copy on demand and shows a green confirmation when it succeeds.

## Updating the app later

Replace the changed files in the repo (Add file → Upload files). The app always fetches the newest files when online, so just close and reopen it.
