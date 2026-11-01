# Licensing (per computer, fully offline)

Every installed copy needs a licence signed by the owner. No server and no internet are involved: the owner holds a
private **owner key**; its public half is built into every copy, which can only *verify* licences, never make them.
This is the same model as the Pharmacy ERP, with the same signing kit (`src/licensing/`, copied unchanged).

## How it works

1. A copy that has no valid licence shows only the **licence screen**, with its **activation code** (for example
   `ABCD-EFGH-JKLM`). Nothing else works: the backend refuses every API call and every report or backup download.
2. The customer sends the code to the owner. The owner issues a licence for that code and sends back the licence text.
3. The customer pastes the licence into the box and presses **Activate**. The normal screens (create account, sign in)
   follow. A licence is for one computer, one person and a set number of days.
4. When 14 days or fewer remain, a small notice in the corner offers **Renew**. A renewal is the same steps with a new
   licence for the same person. If a licence expires while the app is open, the licence screen appears over it; renew and
   you carry on exactly where you were.

## Owner

- **First-time setup (once, on your own computer, from the repo root):** `python src/licensing/make_owner_key.py`
  writes `dic-owner.key` (git-ignored: back it up in two places, never commit, email or paste it anywhere) and fills in
  `src/licence_owner_key.py`; commit that file. Until it is set, the app refuses to run and the release pipeline refuses
  to build.
- **Never run `make_owner_key.py` again once a key is set:** a new key stops every licence already issued.
- **Issue a licence:** `python src/licensing/issue_licence.py --code ABCD-EFGH-JKLM --user "Customer Name" --days 365 --log licences-issued.csv`
  and send back only the licence text. Renewal is the same command with the same name.
- **Your own computer:** on the licence screen choose *I'm the owner of this system* and paste the owner key once. It is
  used in memory and never stored.
- **A licence is refused and you want to know why:** `python src/licensing/verify_licence.py --token <text> --code <their code>`.

## What it does not touch

- The licence lives in its own small file, `licence-state.json`, in the data folder. It is never part of a backup or a
  restore, and it never changes the database, the account, or the sample data.
- Updates keep working on a locked copy, so an expired copy can always update.

## Honest limits

- A licence cannot be revoked. It works until it expires, so keep them short.
- The check is client-side: it stops casual copying and sharing, not someone who edits and rebuilds the code.
- Copying the data folder to another PC copies that copy's install identity.
- The old desktop-window mode (`--desktop`) calls the API directly and bypasses all server-side checks, including the
  sign-in PIN. The shipped installer does not use it.
- The CI pipeline refuses to ship without the owner key, runs the kit self-test, and checks that a fresh copy of the
  built app is locked.

The gate is `src/licence_gate.py`; `src/licensing/` is the vendored kit and must not be reformatted.
Local development needs the owner key too (or a licence for your dev copy).
