# Droid Tycoon Companion

A companion toolkit for **Fortnite: Star Wars Droid Tycoon** — track your droid collection, plan your rebirths, and never accidentally sell something you needed.

## What's in here

```
droid-tycoon-companion/
├── database/
│   ├── droids.json      # All known droids: name, type, rarity, income per tier
│   └── rebirths.json    # All known rebirth levels: credit cost + required droids/tiers
└── excel/
    └── DroidTycoonCompanion.xlsx   # Ready-to-use tracker built from the database
```

## Quick start

Download `excel/DroidTycoonCompanion.xlsx` and open it in Excel or LibreOffice.

1. Go to the **Quick Entry** tab. Pick a droid from the dropdown, set **Owned?** to Yes, then pick the tier you have (Base / Gold / Diamond / Rainbow / Beskar).
2. The **Droid Database** tab updates automatically with your full collection and total income.
3. On the **Dashboard** tab, enter your current credits and the highest rebirth you've already completed.
4. The **Rebirth Tracker** tab shows, level by level, exactly which droids and how many credits you still need — green means ready.
5. **Upgrade Costs** is a quick reference for how many Upgrade Chips each rarity needs per tier.

## The database files

`droids.json` and `rebirths.json` are the source of truth the spreadsheet is built from. Each droid has an `id`, `name`, `type`, `rarity`, and `income` per tier; each rebirth entry lists its credit requirement and the specific droids/tiers needed.

These are meant to be reused — for a future app, a script, or just as a clean reference — without having to re-scrape or retype anything.

## Data sources & accuracy

Droid stats and rebirth requirements are community-sourced (Insider Gaming's Droidex and Rebirth Requirements guides, last cross-checked July 2026). Epic patches this game regularly, so:

- Numbers here may lag behind the live game after a patch.
- A couple of droid income values (`LNG-Shot`, `Proto-Roller`) were adjusted slightly to match the game's consistent tier-doubling pattern — see the `notes` field on those entries in `droids.json`.
- Rebirth levels 18–19 may have shifted slightly in a later patch per player reports — see the `notes` field in `rebirths.json`.

Always cross-check against the in-game Rebirth Station kiosk before committing droids or credits to a rebirth.

## Roadmap

- [x] Droid & rebirth database (JSON)
- [x] Excel tracker
- [ ] GitHub Action to auto-regenerate the Excel workbook from the JSON database
- [ ] Companion app (Flutter)

## Contributing / corrections

If you spot outdated numbers after a patch, update the relevant entry in `database/droids.json` or `database/rebirths.json` and bump the `lastUpdated` field.
