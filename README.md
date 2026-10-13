# Heart Walk Registration

Mobile-friendly sign-up form (full name, phone/Employee ID, shirt size, additional guests)
with a password-protected admin page and Excel download.

| URL | Purpose |
| --- | --- |
| `/` | Public form. **Put this URL in the QR code.** |
| `/admin` | Registration list and totals (password protected) |
| `/admin/export.xlsx` | Excel download |

Admin login: any username, password = `ADMIN_PASSWORD`. If it is unset the admin pages are disabled.

## Run locally
    npm install
    ADMIN_PASSWORD=choose-one npm start     # http://localhost:3000
    npm test

## Deploy (Render, uses `render.yaml`)
1. Render dashboard → New → Blueprint → select this repo.
2. Set `ADMIN_PASSWORD` when prompted. The persistent disk (`/var/data`) keeps the SQLite database across deploys.
3. Your URL is `https://<service-name>.onrender.com/`. Generate the QR code from it.

Any host that runs Node >= 22.13 works; set `DATA_DIR` to a persistent directory.
