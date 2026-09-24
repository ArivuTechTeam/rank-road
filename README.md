# Rank Road – ArivuPro study trackers (January 2027)

Four static web apps + a landing page. No server, no database, no build step needed to deploy.

```
site/                    <- DEPLOY THIS FOLDER
  index.html             landing page (pick your exam)
  foundation/            CA Foundation (3, 5, 7, 9 Jan)
  inter-group1/          CA Inter Group I (2, 4, 6 Jan)
  inter-group2/          CA Inter Group II (8, 10, 12 Jan)
  inter-both/            CA Inter both groups (2–12 Jan)
  manifest.webmanifest   "Add to Home screen" support
  sw.js                  offline support (service worker)
  assets/                app icons
  _headers               Netlify / Cloudflare Pages headers
netlify.toml, vercel.json, .github/workflows/pages.yml   hosting configs
source/                  generators (only needed to CHANGE the plans)
```

## Deploy — pick one

**Netlify (easiest, 2 minutes)**
1. Go to app.netlify.com → Add new site → Deploy manually.
2. Drag the `site` folder onto the page. Done — you get a URL like `rankroad.netlify.app`.
3. Custom domain (e.g. `plan.arivupro.com`): Site settings → Domain management → add domain, then add the CNAME record your DNS provider shows.

**Cloudflare Pages**
Workers & Pages → Create → Pages → Upload assets → upload the `site` folder.

**Vercel**
`npm i -g vercel` then run `vercel --prod` inside this folder (vercel.json points it at `site`).

**GitHub Pages**
Push this whole folder to a GitHub repo (branch `main`), then Settings → Pages → Source: *GitHub Actions*. The included workflow publishes `site/` on every push.

**Your own server / cPanel / S3**
Upload the *contents* of `site/` to the web root (e.g. `public_html/`). Must be served over HTTPS for offline mode and "Add to Home screen" to work.

## How student data works
- Progress (ticks, hours, notes, mock scores, ratings) is saved in the student's own browser (localStorage). Nothing is sent to any server; no login.
- It stays on that phone/laptop. To move devices, students use **Backup → Copy code** and paste it on the new device.
- Clearing browser data erases progress — tell students to copy a backup code occasionally.
- Each app uses its own storage key, so all four can live on the same domain.
- Keep the domain stable: progress is tied to the exact site address.

## Updating the plans later
```
cd source
pip install openpyxl pillow
python build_site.py
```
- Foundation syllabus, hours, dates: `build.py`
- Inter syllabus/chapter weights: `inter_data.py`; Inter calendars, hours, mocks: `gen_inter.py`
- App look and behaviour: `foundation_template.html`, `inter_template.html`

This regenerates `site/*/index.html` and the four Excel timetables in `excel/`.
Then **bump `VERSION` in `site/sw.js`** (e.g. `rankroad-v2`) and redeploy so students' phones pick up the new version.

Changing the schedule after students have started is safe for ticks on days that still exist, but ticks are stored by day + task position, so re-shuffling a day's tasks can move a tick to a different task. Avoid restructuring plans mid-attempt.

## Notes
- Fonts load from Google Fonts; without internet the apps fall back to system fonts.
- Exam dates: ICAI Announcement No. 13-CA(EXAM)/JANUARY/2027 dated 23 Sep 2026.
