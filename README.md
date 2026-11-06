# BillEase — POS & Business Management

Node.js + SQLite (a real relational database stored in one file; no separate DB server to manage).
Roles are enforced **on the server** for every request, not just hidden on screen.

## Run on your own computer
    npm install
    npm start        # http://localhost:3000
    npm test         # 60 automated end-to-end checks

First visit shows **Setup** → create the Owner account. Nothing is pre-filled; add products from the Products page.

## Put it on a host (so phones and the shop computer share one database)
The host must give the app a **persistent disk** (all data lives in the `DATA_DIR` folder).
* **Railway / Render / Fly.io**: deploy this folder (a Dockerfile is included), attach a Volume/Disk mounted at `/data`, set `DATA_DIR=/data`. HTTPS is automatic.
* **A VPS** (about $5/month): install Node 18+, `npm install`, run with `pm2 start server.js`, put Caddy or Nginx in front for HTTPS.
* Free tiers that wipe the disk on restart will **lose your data**. Do not use them for real sales.

Environment variables: `PORT`, `DATA_DIR` (default `./data`), `JWT_SECRET` (optional; auto-generated and stored otherwise).
HTTPS is needed for Bluetooth printing and camera access on phones.

## Roles
| | Owner | Manager | Employee |
|---|---|---|---|
| Billing / POS | yes | yes | yes |
| Change discounts at billing | yes | yes | only if granted (otherwise the product's own discount applies) |
| Add / edit / archive products | yes | only if granted | no |
| Cancel orders | yes | only if granted | no |
| Stock intake, customers, creditors, reports | yes | yes (default) | only if granted |
| Create users | managers + employees | employees only | no |
| Settings, backups, restore | yes | no | no |

The Owner can tick/untick every permission per user (Users → Edit). A manager can only grant permissions they hold themselves.
Orders are never deleted: they are **cancelled** (reason required, stock restored, logged). Users are **deactivated**, which logs them out immediately.

## How live sync works
Employee completes an order → one database transaction (stock checked, unique invoice number issued) → server pushes an event to every signed-in owner/manager screen → toast + PDF link appear without refreshing (tick *Auto-open invoice PDF* on the Dashboard to open it automatically). Each order carries a unique reference from the phone, so a double-tap or network retry can never create a second invoice.

## Backups
Automatic daily backup (last 14 kept) plus manual backup / download / restore under Settings. Restore restarts the server, so the host must auto-restart it (Docker, Railway, Render and pm2 do). Errors are written to `data/error.log`.
