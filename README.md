# StorageScout — Microsoft 365 Storage Savings Audit

Find the SharePoint storage you're paying add-on for and don't need — and the departed-user OneDrives heading for Microsoft's archive meter.
StorageScout runs **read-only inside your own tenant**. It has no server and collects nothing; the report is written to your machine.

---

## This repository: the free Lite build

This is **StorageScout Lite**, free and open source under the MIT License. A run prints the SharePoint pooled-storage position to the console and writes `departed-user-onedrives.csv`; it shows no dollar totals and builds no HTML report. Anything below that mentions the HTML report, dollar totals, Solo or Pro belongs to the paid editions at [storagescout.dev](https://storagescout.dev); that code is not in this script. The `-Lite` switch changes nothing in this build: it always runs as Lite.

## What you get

| File | Purpose |
|------|---------|
| `StorageScout-Audit.ps1` | The audit engine (PowerShell + Microsoft Graph) |
| `test/mock-tenant.json` | Demo data to try the tool with no tenant access |

Licensed MIT — Solo and Pro are under the EULA.

Outputs per run: a polished **HTML savings report** + **CSV exports** for each finding (Solo/Pro). Lite: console storage position + `departed-user-onedrives.csv`.

---

## Requirements

- **PowerShell 7+** (Windows, macOS, or Linux) — or Windows PowerShell 5.1
- **Microsoft Graph PowerShell SDK**:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```
- An account that can consent to **read-only** Graph scopes. To pull the storage usage reports you need **Reports.Read.All** (a Global Reader or Reports Reader can consent). `Sites.Read.All` helps enrich site detail.

> Note: unlike license auditing, tenant-wide storage reports are richer when run by an admin with reporting rights. Without report access, StorageScout degrades to a limited view and tells you so.

---

## Quick start (try it with demo data — no tenant needed)

```powershell
.\StorageScout-Audit.ps1 -MockDataPath .\test\mock-tenant.json
```
Open the generated `StorageScout-Report-<timestamp>\StorageScout-Report.html`.

## Run against your tenant

```powershell
# 1) (first time) install the SDK
Install-Module Microsoft.Graph -Scope CurrentUser

# 2) run the audit (you'll be prompted to sign in & consent to read-only scopes)
.\StorageScout-Audit.ps1 -StaleDays 365
```

You'll be asked to consent to: `Reports.Read.All`, `Sites.Read.All`, `User.Read.All`, `Organization.Read.All`, `Directory.Read.All`. All read-only.

---

## Options

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-StaleDays` | `365` | A site with no activity in this many days is flagged stale (archive candidate) |
| `-AddOnPurchasedGB` | auto | GB of Extra File Storage you currently buy. If omitted, your over-pool overage is used |
| `-OutputPath` | auto | Folder for the report + CSVs |
| `-PriceTablePath` | `.\config\StoragePriceTable.csv` | Your add-on storage price |
| `-CompanyName` | — | **(Pro)** White-label the report header |
| `-CurrencySymbol` / `-CurrencyLabel` | `$` / `USD` | Display currency |
| `-Lite` | off | Free mode: pool position + departed-user-OneDrive CSV only (no $ totals, no HTML) |
| `-MockDataPath` | — | Use demo JSON instead of connecting to Graph |

**Example (consultant, white-labelled, 2-year stale threshold):**
```powershell
.\StorageScout-Audit.ps1 -StaleDays 730 -CompanyName "Contoso IT Partners"
```

---

## Set your price (important)

Add-on storage pricing varies by agreement, region and model (Microsoft's public list is ~$0.20/GB/month for Extra File Storage; a pay-as-you-go model entered preview in 2026). Open `config/StoragePriceTable.csv` and set `AddOnPricePerGBMonthly` to your real rate. Everything in the report scales from it.

---

## What it checks

1. **SharePoint pool position** — site storage used vs pooled allowance (1 TB + 10 GB/licensed user) vs add-on purchased. OneDrive is per-user storage and is **not** part of the pool.
2. **Idle add-on** — add-on you buy above your actual overage; cancellable at renewal with no work
3. **Version-history bloat** — reclaimable in place by trimming excess file versions
4. **Recycle-bin bloat** — reclaimable in place by emptying first/second-stage bins
5. **Departed-user OneDrives** — owner deleted, unlicensed or disabled. Deleted/unlicensed accounts follow Microsoft's lifecycle: read-only at day 60, archived at day 93, **deleted at day 365** unless pay-as-you-go billing is enabled (then $0.05/GB/month on every archived account tenant-wide, $0.60/GB to reactivate one). Reported as exposure with the per-account size so you can relicense, move or delete before day 93. Disabled-but-licensed accounts are flagged as license waste.
6. **Stale sites** *(advisory)* — no activity in > N days; archive candidates
7. **Near-quota sites** *(advisory)* — approaching their limit; operational risk

Headline = **add-on storage you can stop paying for** = idle add-on + add-on freed by in-place reclaim (version history + recycle bins), never more than you buy today. Departed-user OneDrives and advisory items are listed but not added to the total.

---

## Security & privacy

- **Read-only.** StorageScout never writes or changes anything in your tenant.
- **No exfiltration.** No network calls except to Microsoft Graph. The report is generated locally.
- Review the script before running — it's plain PowerShell, intentionally readable.
- Always honor retention / legal-hold before reclaiming any storage or deleting a departed user's OneDrive.

---

## Troubleshooting

- *"Usage reports limited"* → the account lacks `Reports.Read.All`; results still valid but partial, real reclaimable storage is likely higher.
- *Every departed-user status shows `deleted`* → the account lacks `User.Read.All`, so owner lookups fail; grant it to distinguish deleted / unlicensed / disabled.
- *Module not found* → `Install-Module Microsoft.Graph -Scope CurrentUser`
- *Report shows names as opaque IDs* → your tenant has "conceal user names in reports" enabled (M365 admin center → Settings → Org settings → Reports). Turn it off (or accept anonymized output).
- *A price looks wrong* → set `AddOnPricePerGBMonthly` in `config/StoragePriceTable.csv`.

---

StorageScout · storagescout.dev · Independent tool, not affiliated with or endorsed by Microsoft.
Part of the Scout suite · pairs with SeatScout (license waste) — seatscout.dev

## Author

Built and maintained by Uğur Sayar, a SharePoint and Microsoft 365 architect in Istanbul ([LinkedIn](https://www.linkedin.com/in/ugursayar)). Bug reports and pull requests are welcome through GitHub issues.
