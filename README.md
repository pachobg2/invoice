# invoice

Windows invoice and warehouse-receipt program (Bulgarian, EUR) with a portable edition for a flash drive.
The program contains **no company, customer or invoice data** — it starts empty. Data lives only in the
`invoice-data` folder next to `invoice.exe`.

| Part | What it is |
|---|---|
| `invoice.exe` | Main app: .NET 8 WinForms + WebView2 host, web UI from `faktura/` embedded in the exe |
| `invoice-lite.exe` | Portable edition for flash drives, .NET 2.0, runs on Windows XP – 11 (built into `invoice.exe`, which copies it to the flash drive) |

## Updates

`invoice.exe` checks the latest GitHub release of this repository (about once a day, can be turned off in
Настройки → За програмата), downloads `invoice.exe`, verifies it against `SHA256SUMS.txt` from the same
release and restarts itself. `invoice-data` is never touched. The Lite edition on a flash drive is refreshed
from the main app during "sync from flash".

## Build

```powershell
powershell -ExecutionPolicy Bypass -File build.ps1      # builds Lite, then invoice.exe (needs the .NET 8 SDK)
```

## Release

1. Bump `<Version>` in `host/Faktura.csproj` (and `lite/Faktura-Lite.csproj` + `LITE_VERSION` in `faktura/app.js` if Lite changed)
   and add a matching first entry to `CHANGELOG` in `faktura/app.js`.
2. Commit, then `powershell -ExecutionPolicy Bypass -File release.ps1` — builds, tags `vX.Y.Z`, pushes and
   creates the GitHub release with `invoice.exe`, `invoice-lite.exe` and `SHA256SUMS.txt`.

No license is granted; all rights reserved.
