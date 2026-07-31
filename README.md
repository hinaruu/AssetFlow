# Asset Manager

A shared, online-only asset tracking app. Everyone who signs in reads and
writes the **same** data — there's one location, one shared dataset, and
you (as Admin) see every change anyone makes. There is no offline mode:
if there's no internet, the app tells you instead of silently failing.

## How it works

- Data lives in a small **Supabase** database (free tier) — not in the app itself.
- When the app loads, it fetches the current shared data.
- When anyone saves a change, it's written straight back to that same database.
- The **🔄 Sync** button in the top bar re-fetches the latest data on demand,
  in case someone else made changes elsewhere.
- Login is username/password (kept in the shared data), with **Admin** vs
  **Regional Staff** roles already built in — Admin sees and edits everything.

## 1. Create the database (5 minutes, free)

1. Go to [supabase.com](https://supabase.com) → sign up → **New project**
2. Once it's created, open **SQL Editor** → **New query**
3. Paste in the contents of `supabase-setup.sql` (included in this project) and click **Run**
4. Go to **Project Settings → API** and copy two values:
   - **Project URL**
   - **anon public** key

## 2. Connect the app to it

Copy `.env.example` to a new file named `.env`, and fill in the two values from step 1:

```
VITE_SUPABASE_URL=https://your-project-ref.supabase.co
VITE_SUPABASE_ANON_KEY=your-anon-public-key
```

Then run it locally:

```bash
npm install
npm run dev
```

The first time it loads, it seeds the database with demo data automatically.

**Demo logins:**
- Admin — `admin` / `admin123`
- Staff — `staff` / `hr12345`

(Change these in the **Users** page once you're in as Admin.)

## 3. Put it on GitHub Pages

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO-NAME.git
git push -u origin main
```

Then, since `.env` is never pushed to GitHub (it holds your database key),
give GitHub Actions the same two values as **secrets**:

1. Repo → **Settings → Secrets and variables → Actions → New repository secret**
2. Add `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY` with the same values as your `.env`
3. Repo → **Settings → Pages** → set **Source** to **GitHub Actions**
4. Push to `main` (or re-run the workflow under the **Actions** tab)

Your app goes live at `https://YOUR-USERNAME.github.io/YOUR-REPO-NAME/`.
The included workflow (`.github/workflows/deploy.yml`) rebuilds and redeploys
automatically on every push to `main`.

## 4. Install it on your phone

**iPhone (Safari):** open the URL → Share icon → **Add to Home Screen**
**Android (Chrome):** open the URL → **⋮** menu → **Install app**
**Desktop (Chrome/Edge):** click the install icon (⊕) in the address bar

It'll open like a normal app, with its own icon — but it always needs
internet to actually load or save data, by design.

## Notes on security

For simplicity, the database table is set up so anyone with your site's
URL and public key can read/write it — access control happens at the
app's login screen, not the database. That's fine for a small internal
tool, but don't store anything sensitive (passwords, financial data,
personal info you wouldn't want exposed) in it. If you need real
per-user security later, Claude can help add Supabase Auth and
row-level permissions.
