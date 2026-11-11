# Lift Board

A local dashboard that reads a lifting log from a Google Sheet and tells you when to move up in weight.

## What it does

- **Move-up calls.** For every lift on every training day it predicts how many reps the next weight will cost (Epley) and says whether to load it, keep building, hold, or drop back.
- **Muscle maps.** Front and back body figures for each exercise, each day and the whole week, split into primary and secondary muscles.
- **Weekly volume.** Sets or total pounds per week, by muscle and by lift, with one page per training day.
- **Death sets.** Drop chains from a tracker tab, week over week.
- **Time in the gym.** Session length, pace (minutes per set or pounds per minute) and change versus the last time on that day.
- **Daily summary.** A shareable PNG card per workout with volume, % change, muscles hit and the lift table.

## Sheet format

- One tab per training day. The header row lists the lifts as `Exercise 4xf (S4 death)`; a later header row on the same tab starts a new routine.
- Column A is the date. Each lift cell is a formula `weight*(reps+reps)+weight*(reps)`, for example `80*(8+6)+70*(9)`.
- Cyan fill marks a death-set day; blue fill marks a different machine (left out of the calls).
- An optional `Time` column holds the session length as a duration.
- A tab named `Attempt at death set tracker` holds weekly drop chains, with columns `Day - Lift`.

## Setup

1. Share the sheet by link, then put its ID (the long code in its URL) in `sheet_id.txt`.
2. Start the server: `python3 server.py`, then open http://localhost:8765.
3. Click **Refresh from sheet** after a workout. It downloads the sheet into `data/`, rebuilds `index.html` and reloads.

Settings live at the top of `build.py`: rep floors, the rep cap, machine base weights (`BASE_WEIGHT`), lift name aliases and the muscle map (`MUSCLES`).

`shots.mjs` saves screenshots of each board section and every summary card (`node --experimental-websocket shots.mjs`, with Chrome installed and the server running).

No dependencies beyond Python 3.9+ standard library.
