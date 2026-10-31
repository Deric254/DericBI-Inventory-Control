# Data, backups and recovery

DericBI Inventory Control keeps all data in one **SQLite** database (`inventory.db`) in the app's data folder
(`%APPDATA%\DericBI Inventory Control` on Windows). Every change is a transaction: it is saved completely or not at
all, even if the PC loses power in the middle of it.

## Automatic protection
| What | When | Where (`backups\`) | Kept |
|---|---|---|---|
| Rolling snapshot | before changes, at most every 5 min | `ICT_backup_*.db` | newest 10 |
| Daily snapshot | first change each day | `daily\ICT_daily_*.db` | 30 days |
| Safety copy | before every delete, replacement or restore | `pre-restore\` | newest 10 |
| Pre-update copy | before an app update is installed | `pre-update\` | newest 3 |
| Original Excel file | once, when upgrading from the Excel version | `pre-sqlite\` + `ICT_MASTER.pre-sqlite.xlsx` | forever |

## Everyday tasks
* **Excel copy for reports/sharing:** sidebar → *Export to Excel*. This creates a fresh `.xlsx` in `exports\` (newest 5 kept).
  Editing that file does not change the app's data.
* **Move to a new PC / full backup:** Security → *Export Full Backup*, then *Restore* on the new PC. The zip contains the
  database, an Excel copy, your PIN file and settings.
* **Undo a mistake:** Security → Backups, or restore a file from `backups\daily\`.
* **Create your account:** on first launch you choose a username and a PIN. Installs made before accounts existed keep
  signing in with the PIN alone; nothing is forced on them. Forgot your username? Enter your recovery key on the sign-in
  screen and it is shown to you.
* **Start with your own data:** a fresh install comes with a few sample assets (serial numbers starting `SAMPLE-`).
  Security → *Sample Data* → *Remove sample data* deletes only those, after saving a safety copy. The card disappears once
  they are gone, and they never come back.

## Rules the database enforces
* Serial numbers are unique among active assets (blank / "N/A" / "none" are allowed to repeat; Disposed and Replaced assets
  free their serial for reuse).
* Equipment Type is required. A transfer must change the department. Ink cannot be used or moved beyond available stock.
* Asset ids never change and are never reused, so a screen that is open on an old list can never edit the wrong asset.

## Limits (be aware)
* The PIN protects the **application**. The database file itself is not encrypted: use Windows BitLocker on computers that
  hold sensitive data.
* The interface loads the full asset list into memory. Comfortable up to roughly 100,000 assets (about 40 MB per refresh);
  every save is still instant at any size.
* Keep the data folder on a local drive (not a network share or a sync folder).
