# Deploying on the office server (for IT)

The app is a single Python process plus a SQLite database file. It runs on Windows or Linux.
It must sit behind a reverse proxy that terminates HTTPS; the app itself only listens on
`127.0.0.1`.

## What IT sets up

| Item | Notes |
|---|---|
| Python 3.11+ | `py -3.12 -m venv C:\kns-plan\venv` |
| LibreOffice | Used headless to recalculate and check every workbook before download. About 350 MB. If it is not in the default folder, set `KNS_SOFFICE` to `soffice.exe`. |
| Reverse proxy with HTTPS | IIS (URL Rewrite + ARR), nginx or Caddy, with the office domain's certificate. Forward to `http://127.0.0.1:8400` and set `X-Forwarded-For`. |
| Access from outside | Strongly recommended: VPN or an IP allow-list on the proxy. The app has one shared password; it locks an IP out for 15 minutes after 5 wrong attempts. |
| Backups | Nightly scheduled task (below), to a different disk or a network share. |

## Install

```
py -3.12 -m venv C:\kns-plan\venv
C:\kns-plan\venv\Scripts\pip install "C:\path\to\KNS-Business-Plan-Generator[web]"
mkdir D:\kns-plan-data
C:\kns-plan\venv\Scripts\kns-plan hash-password
```

Save the printed hash as `D:\kns-plan-data\password.hash` (or as the `KNS_PASSWORD_HASH`
environment variable). The password itself is never stored.

Set the company default cost stack (rates new plans start with). It lives only on the server:

```
C:\kns-plan\venv\Scripts\kns-plan set-default-stack plan-export.json --data D:\kns-plan-data
```

`plan-export.json` is any plan's data downloaded from the app's Review step; the first cost stack
is used (or pass `--stack <ID>`). Finance can provide the file.

## Run as a service

Environment for the service:

```
KNS_DATA_DIR=D:\kns-plan-data
KNS_HTTPS=1
```

Command:

```
C:\kns-plan\venv\Scripts\kns-plan serve --host 127.0.0.1 --port 8400
```

On Windows use NSSM or Task Scheduler ("At startup", "Run whether user is logged on or not").
Health check: `GET http://127.0.0.1:8400/health` returns `{"ok": true}`.

Generation runs one workbook at a time (LibreOffice is single-instance); a plan takes a few seconds.

## Nightly backup

```
C:\kns-plan\venv\Scripts\kns-plan backup --data D:\kns-plan-data --to \\nas\backups\kns-plan --keep 30
```

This takes a consistent copy of `plans.sqlite` (every version of every plan). Generated workbooks
in `D:\kns-plan-data\outputs` can be regenerated from any version, so backing them up is optional.

## Changing the password

Run `kns-plan hash-password` again, replace `password.hash` and restart the service. Everyone
signs in again.

## Data folder contents

| Path | What |
|---|---|
| `plans.sqlite` | Every plan and every saved version; login attempts |
| `outputs/<plan>/` | Generated workbooks and their check reports |
| `uploads/` | Uploaded files waiting for column matching (deleted after import) |
| `.session_key` | Signs session cookies. Deleting it signs everyone out. |
| `password.hash` | Shared password hash (if not set as an environment variable) |
| `default_stack.json` | The company default cost stack for new plans |
