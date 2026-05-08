# HRV Screenshot Uploader

Static React app for uploading HRV screenshots, running local OCR with Tesseract.js, reviewing extracted values, and posting confirmed values to Google Sheet through Apps Script.

The app also evaluates each upload with the rules in [HRV_INTERPRETATION_SPEC.md](./HRV_INTERPRETATION_SPEC.md), including single-metric bands, comprehensive diagnosis, recommendations, and a browser-local 7-record baseline.

## Run locally

```bash
python3 -m http.server 4173
```

Then open:

```text
http://127.0.0.1:4173/
```

The page loads React, lucide-react, framer-motion, Tailwind, and Tesseract.js from CDNs.
