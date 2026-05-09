# HRV Screenshot Uploader

Static React app for uploading HRV screenshots, running local OCR with Tesseract.js, reviewing extracted values, and posting confirmed values to Google Sheet through Apps Script.

The app also evaluates each upload with the rules in [HRV_INTERPRETATION_SPEC.md](./HRV_INTERPRETATION_SPEC.md), including single-metric bands, comprehensive diagnosis, recommendations, BSRS mood thermometer scoring, patient profiles, history trend lookup, Google Sheet history lookup, and a browser-local 7-record baseline.

## Google Sheet history API

Use [APPS_SCRIPT_WEB_APP.gs](./APPS_SCRIPT_WEB_APP.gs) as the Apps Script web app code, then deploy it as a web app accessible to anyone.

Deploy settings:

- Execute as: `Me`
- Who has access: `Anyone`
- After every code change, create a new deployment or edit the existing deployment version.

The frontend writes records with `POST` and reads cloud history with:

```text
GET ?action=history&email=user@example.com&callback=...
```

## Run locally

```bash
python3 -m http.server 4173
```

Then open:

```text
http://127.0.0.1:4173/
```

The page loads React, lucide-react, framer-motion, Tailwind, and Tesseract.js from CDNs.
