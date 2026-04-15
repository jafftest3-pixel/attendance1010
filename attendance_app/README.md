# Attendance System (Netlify)

This folder contains a **browser-based** attendance app that deploys to **Netlify**. Data is stored in **IndexedDB** on each user’s device (not on a central server).

## Deploy on Netlify

1. Push the `attendance_app` folder (or repo with this as root) to GitHub.
2. In Netlify: **Add new site → Import an existing project**.
3. Set **Publish directory** to `public` (or rely on `netlify.toml` if the repo root is `attendance_app`).
4. Deploy.

The **Google Sheet sync** features call `/.netlify/functions/sheet-proxy`, which only runs after deploy (or locally with Netlify CLI).

### Local preview with functions

```powershell
cd C:\Users\B0008\Desktop\attendance_app
npx netlify-cli dev
```

Then open the URL Netlify prints (Sheet sync will work).

## Legacy Flask app

Files such as `app.py`, `wsgi.py`, and `templates/` are the original **Python/Flask** version. They are **not** used by the Netlify static app in `public/`. To run Flask locally:

```powershell
pip install -r requirements.txt
python app.py
```

The Flask app now exposes a shared settings API at `/api/settings`. When the static app is served from the same backend host, Google Sheet links and other saved settings are saved centrally and become visible to all users of that instance.

## Netlify app layout

- `public/index.html` — single-page UI
- `public/js/storage.js` — IndexedDB (Dexie)
- `public/js/main.js` — logic, imports, reports
- `public/styles.css` — styles
- `netlify/functions/sheet-proxy.js` — serverless proxy for Google Sheets CSV
