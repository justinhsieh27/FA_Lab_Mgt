# FA Lab Case Management

An internal web system for a semiconductor FA Lab. It uses one Excel workbook as the operational record with exactly two main sheets:

- `FA Case` — requester-facing case data
- `Job` — sample/cut-level execution data and step history

## First start

```bash
./install.sh
./run.sh
```

`run.sh` starts the service in the background, storing its PID in `fa-lab.pid` and its output in `logs/fa-lab.log`. To use another port, run `PORT=3001 ./run.sh`.

Open `http://localhost:3000` on the host computer. Other users on the same internal network can use `http://<FA-Lab-host-IP>:3000`.

Initial accounts (change the passwords before production use):

| Username | Password | Role |
| --- | --- | --- |
| `manager` | `ChangeMe!2026` | Manager |
| `engineer01` … `engineer08` | `ChangeMe!2026` | Engineer |
| `guest` | `ChangeMe!2026` | Guest |

Use **User Access** as Manager to create, update, deactivate, or reset accounts. The Manager can update their own password from this screen as well.

## Data and backup locations

- Main workbook: `data/FA_Lab_Case_Management.xlsx`
- Account records and immutable audit history: `data/system.sqlite`
- Uploads: `uploads/`
- Daily backups: `backups/<timestamp>/`

The service makes a backup at startup and once every 24 hours. Keep these folders on a managed FA Lab drive and include the project folder in the site backup policy.

## Workflow rules

- A Case No. is entered manually and must be unique.
- New cases default to `PRS` and `Normal` priority.
- Batch Job creation accepts `Sample ID | Location | Cut count` on one line, such as `S1|Center|2`; it produces `<Case No.>.S1C1`, `<Case No.>.S1C2`, etc.
- Each Job starts as `Sample preparation → FIB or SELA → TEM or SEM`.
- Steps are sequential by default. A Manager can enable a documented sequence override and add custom steps.
- The Case status derives from its active Jobs. Manager-only `On Hold` and `Cancelled` overrides take priority.
- Cancelled Jobs do not affect Case completion or Case `end_date`.
- A completed Case `end_date` is the average of valid completed Job end dates. Cycle time uses calendar days.

## Deployment note

This first version intentionally runs as a local internal service. Before exposing it beyond the FAB intranet, place it behind the factory's approved HTTPS reverse proxy and firewall policy. Do not expose port 3000 directly to the public Internet.
