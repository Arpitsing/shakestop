# ShakeStop 🥤

Concept prototype for **ShakeStop** — a protein‑shake vending machine business.

**Live site:** `https://<your-username>.github.io/shakestop/`

## Pages
- **index.html** — the **single scrollable site**: hero, problem, solution, how-it-works, the **live interactive machine (embedded)**, effort split, 5% revenue model, why-partner, 30-day pilot and contact — plus a **Download .pptx** button.
- **shakestop-prototype.html** — the interactive vending machine prototype (embedded into `index.html`, also opens standalone).
- **shakestop-logo.html** — the ShakeStop logo system (primary, stacked, icon, monochrome, palette).
- **ShakeStop-Gym-Pitch.pptx** — the downloadable PowerPoint version of the pitch.

All pages are self‑contained static HTML — no build step, no dependencies.

---

## How to host it free on GitHub Pages

### Option A — GitHub website only (no tools, easiest)
1. Sign in at <https://github.com> (create a free account if needed).
2. Click **+** (top‑right) → **New repository**.
   - **Repository name:** `shakestop`
   - **Public** (required for free Pages)
   - Tick **Add a README** is optional — you already have one.
   - Click **Create repository**.
3. On the new repo page click **Add file → Upload files**.
4. Drag in **all files**: `index.html`, `shakestop-prototype.html`, `shakestop-logo.html`, `ShakeStop-Gym-Pitch.pptx`, `README.md`.
5. Click **Commit changes**.
6. Go to **Settings → Pages** (left sidebar).
7. Under **Build and deployment → Source**, choose **Deploy from a branch**.
8. Set **Branch = `main`**, **Folder = `/ (root)`**, click **Save**.
9. Wait ~1 minute, refresh. GitHub shows:
   **“Your site is live at `https://<your-username>.github.io/shakestop/`”** ✅

### Option B — Git command line
```bash
# in this folder
git init
git add index.html shakestop-prototype.html shakestop-logo.html ShakeStop-Gym-Pitch.pptx README.md .gitignore
git commit -m "ShakeStop prototype site"
git branch -M main
git remote add origin https://github.com/<your-username>/shakestop.git
git push -u origin main
```
Then do **Settings → Pages** steps 6–9 above.

---

## Updating the site later
Edit any file, then **upload again** (Option A) or:
```bash
git add -A
git commit -m "Update"
git push
```
Changes go live in about a minute.

## Notes
- Keep the repository **Public** — private repos need a paid plan for Pages.
- The URL is case‑sensitive and matches your repo name.
- Brand names / logos shown (ON, MuscleBlaze, Anytime Fitness) are **placeholders** for a
  concept mock; use only licensed assets before any commercial launch.
