# The dericBI ecosystem: how your systems know about each other

Goal: your systems (CRM, website and Health Check, Business Reality Engine, and the business systems) work as one family without any of them
being written around another. No system's code names another system. Each one only (1) publishes a small **capability card** and (2) keeps
a list of addresses. A new system is a new address; a system that is off just shows "not reachable".

## The capability card (contract `dericbi.ecosystem/1`)
A small JSON file every system publishes. Only `contract`, `id` and `name` are required.

```json
{ "contract": "dericbi.ecosystem/1", "id": "bre", "name": "Business Reality Engine", "version": "1.0",
  "role": "what it is, in a few words",
  "offers": ["alerts", "forecast"], "accepts": ["events.webhook", "files.upload"], "health": "/api/health/" }
```
`offers` and `accepts` are plain strings. Any system may invent new ones; nobody needs a central list. Unknown fields are kept and ignored.

A card may also carry `"handoffs": {"licence.issue": "/master?issue_to={name}&ref={ref}"}`: a ready-made link (a plain path on that same system, `{name}` and `{ref}` filled in by the opener) that another system can show its owner. Anything that is not a plain path is dropped. Opening it sends nothing; it only fills a form that the owner still submits. The CRM shows it as a button on a won lead (`{ref}` is the lead's permanent reference, the same on any copy of the CRM database, kept on the licence only so alerts can be matched back).
A card may also list `"feeds": {"advisor.alerts": "/api/advisor/alerts"}`: a place another system may read from, under that system's own rules.
Cards hold no secrets and no business data.

## Where each card lives (the CRM tries these three paths in order)
1. `/.well-known/dericbi.json`   2. `/dericbi.json`   3. `/api/ecosystem/card`

| System | Kind | Card at | Reachable at |
|---|---|---|---|
| Website and Health Check | static site | `/.well-known/dericbi.json`, `/dericbi.json` | the public site |
| Portfolio | static site | `/.well-known/dericbi.json`, `/dericbi.json` | the public site |
| Business Reality Engine | hosted, many businesses | `/api/ecosystem/card` | its own address; no database needed for the card |
| CRM | desktop (yours only) | `/api/ecosystem/card` | needs sign-in; for display |
| Inventory Control | desktop, per customer | `/api/ecosystem/card` | `http://127.0.0.1:47615` while it runs (any free port only if that one is taken) |
| Imara Links | desktop, per customer | `/api/ecosystem/card` | `http://127.0.0.1:9000` while it runs (any free port only if that one is taken) |
| smePro | desktop and phone, per customer | none yet | `http://127.0.0.1:8080` while it runs |
| Pharmacy ERP | desktop, per customer | none, on purpose | a fresh free port each start, chosen so it can never clash with another app |

The cards are the truth; this table can lag behind them.

The desktop programs that publish a card serve it from a plain data file (Inventory Control: `src/ecosystem.json`; Imara Links: `backend/ecosystem.json`). Edit the file to change what the card says, no code. If the file is missing or broken the address answers 404 and nothing else is affected. The card is served without sign-in and without a licence, because it holds no data and other systems need to find the app before it is licensed.

A desktop program has no public web address, so it can only be reached while it is running, on the same computer. That is why a card only matters for the programs on your own machine. Customers' copies run on their own computers and are never reached from outside; they stay fully independent.

## In the CRM
Settings > Data and backups > **Your other systems**: add a name and address, press Add and check. It shows reachable or not, what the
system is, and what it offers and accepts. Owner only. Nothing is sent to those systems from the CRM.

## Rules for adding the next system
- Publish a card. Do not edit another system to know about yours.
- Capabilities are strings; add one when you add a feature.
- A system that is missing, down, or on a newer card version must never break another: show it in words and carry on.
- When data starts to flow (next step), it goes through each system's own existing door (for example BRE's per-connector webhook token) with
  a retry queue, so a down system never loses or blocks anything.

## Alerts from businesses back to the CRM (opt-in, narrow)
Many businesses use the Business Reality Engine; their alerts are theirs. So nothing flows by default:
1. A business only counts if its licence was issued from a CRM lead (the licence carries that lead's number).
2. Its own owner must switch on **Tell my DericBI advisor about serious problems** (Settings; off by default, can be turned off at any time, and says what is shared).
3. Only critical and high alerts that are still open, and only: lead number, business name, severity, the headline, the time. Never records, figures behind it, or explanations. A headline can name a branch or product and contain a percentage ("Westlands sales are down 15%"); the switch says so.
4. Read-only, behind a key the CRM makes up by itself. Nothing is set up by hand: when you open the engine from a won lead, the link carries that key after the `#` (never sent to a server or logged). The engine, once you are signed in as the platform owner, keeps only a hash of it. Until that has happened the feed does not exist (404) and the CRM says "Not connected yet".
5. The CRM asks every 5 minutes (it is a desktop app, so the engine cannot call it). Each alert becomes one follow-up on that lead, once: due today unless one is already due or overdue, and a later planned date that is brought forward is written in the lead's history. "Check alerts now" does it on demand and says in words what happened.

## Not built yet (on purpose)
1. Quiet clients (no data for days) coming back as follow-ups, and anything flowing the other way.
2. Shared fixture files so the Health Check page and BRE read messy files identically.
Each will be off by default and use the cards to decide whether it applies.
