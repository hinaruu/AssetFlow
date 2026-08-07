# Critical Security Fixes — Step by Step

This walks through fixing the six 🔴 Critical items from the architecture
review: open database policies, fake password hashing, client-only
permission checks, shipped default credentials, and a tamperable audit
log. It assumes no prior Supabase Auth experience.

**I've already made the code changes** described in Phase 3 below — they're
in this zip. You mainly need to do the Supabase-dashboard steps (Phases
1, 2, 4, 6) and then test (Phase 7).

Do these phases **in this order** — some depend on the last (in
particular: deploy the new code in Phase 3 *before* creating your first
account in Phase 4 — the app can only auto-seed your locations/categories
the very first time it loads with a signed-in session, and if a profile
row already exists that trick won't fire).

---

## Before you start

- **Back up your current data first.** In the app, go to **Backup & Restore**
  (as Admin) → **Download Backup**. Keep that file somewhere safe. If
  anything goes wrong below, this is how you'd recover.
- **If you can, test this on a second (free) Supabase project first**,
  rather than your real one — create a new project at supabase.com, run
  through everything below there, confirm it works, *then* repeat on your
  real project. If that's not practical, at least do this during a quiet
  period, since logins will be briefly interrupted (between Phase 2 and
  Phase 3).

---

## Phase 1 — Turn on email login in Supabase

1. Open your project at [supabase.com](https://supabase.com) → **Authentication** (left sidebar) → **Providers**.
2. Confirm **Email** is enabled (it is by default).
3. Click **Email** to expand it, and for an internal company tool, turn
   **Confirm email** **off** — otherwise every new person needs to click a
   confirmation link before they can log in, which adds friction for
   accounts you're creating for people yourself in Phase 4/6. (If you'd
   rather keep email confirmation on, that's fine too — just know new
   accounts won't be able to sign in until they confirm.)
4. Save.

This alone doesn't change your app yet — it just turns on the login system
you're about to use.

---

## Phase 2 — Run the database migration

1. In Supabase, go to **SQL Editor** → **New query**.
2. Open `critical-security-migration.sql` (included in this zip), copy
   its entire contents, paste into the editor, and click **Run**.
3. You should see "Success. No rows returned." If you get an error, stop
   and re-read the error message — the most likely cause is that an
   earlier migration (like `add-pending-deletion-column.sql`) was never
   run against this project. Run any of the older migration `.sql` files
   in this folder you haven't run yet, then retry this one.

What this did: added a column linking your `users` table to real Supabase
logins, and replaced every "anyone can do anything" database policy with
one scoped to Admin / Regional Admin / Regional Staff and their location.

**Your app's old login screen will stop working correctly right after
this** (it compares passwords against a `users` table it can no longer
freely read) — that's expected. Move on to Phase 3 right away.

---

## Phase 3 — Deploy the updated app code

The code in this zip already includes the fix — here's what changed, so
you know what you're deploying:

- **`src/lib/supabase.js`** — the `users` table no longer sends/reads a
  password hash; it reads/writes `authUserId` (linked to the Supabase
  login) instead.
- **`src/App.jsx`**:
  - The login screen now calls real Supabase Auth (`supabase.auth.signInWithPassword`)
    instead of comparing a hash in the browser. It asks for **email**
    now, not username.
  - The app tracks who's logged in via the actual Auth session
    (`supabase.auth.onAuthStateChange`), not a plain piece of React
    state — refreshing the page keeps you logged in now, and signing out
    (`supabase.auth.signOut()`) is a real thing rather than just clearing
    local state.
  - Company data (assets, locations, etc.) is now only fetched **after**
    you're signed in — the database no longer allows anonymous reads, so
    there's nothing to fetch before that. You'll see the login screen
    immediately on load, then a brief spinner right after you sign in.
  - If someone signs in but has no matching profile in `users` yet, they
    see a clear "Account not set up yet" screen instead of a confusing
    blank app.
  - The demo `admin123` / `hr12345` accounts are gone.
  - **User Accounts** page: "New User" now only creates a profile — it
    tells the Admin to invite that email in Supabase and link it (Phase 6).
    "Reset Password" now sends a real password-reset email instead of
    setting a fake hash.

To deploy: build and publish this zip's contents the same way you did
before (e.g. `npm install && npm run build`, then push to GitHub Pages /
wherever this is hosted).

Don't sign in yet — go to Phase 4 first, since the very next sign-in is
what triggers the one-time seed of your locations/categories.

---

## Phase 4 — Create your first real Admin account

Because creating a login has to happen with elevated (`service_role`)
access that must never live in the app's code, the very first account has
to be created by hand, in the dashboard, once. (Creating *further*
accounts gets easier — see Phase 6.)

1. In Supabase: **Authentication** → **Users** → **Add user** → **Create new user**.
2. Enter your own email and a strong password. Leave "Auto Confirm User"
   checked. Click **Create user**.
3. **Go to the live app now and sign in** with that email/password. You'll
   land on the "Account not set up yet" screen — that's expected, you
   haven't linked a profile yet. What just happened behind the scenes:
   because every table was still completely empty (including `users`),
   signing in triggered the app's one-time seed of demo locations and
   categories. This only happens once, on a truly empty database — which
   is why the order here matters.
4. Back in Supabase: **Authentication** → **Users**, click into the
   account you just created, and **copy their User UID** (a long string
   like `3fa85f64-5717-4562-b3fc-2c963f66afa6`).
5. **SQL Editor** → **New query**, run this (replace both placeholders —
   the email must exactly match what you signed in with):

   ```sql
   insert into users (id, name, username, email, role, location_id, auth_user_id)
   values (
     'usr-admin-1',
     'Your Name',
     'yourname',
     'you@company.com',           -- must match the email from step 2
     'Admin',
     null,
     'PASTE-THE-USER-UID-HERE'
   );
   ```

6. Refresh the app. You should now land inside as Admin, with the seeded
   demo locations/categories visible.

---

## Phase 5 — Confirm the audit log can't be edited

Quick sanity check that the "no update/delete" policy from Phase 2 is
actually in effect: in Supabase, **Table Editor** → `audit_log`, note that
Phase 2 ran without error when it created a policy set with no update/
delete rule for this table — that absence *is* the protection (with RLS
on and no policy for an action, that action is refused for every normal
client, including your own app). You don't need to do anything further
here; this is really just a note on what to expect if you ever go looking
for an "edit" or "delete" option on this table from the app and don't
find one — that's by design now, not a bug.

---

## Phase 6 — Add your remaining users

For each additional person (Regional Admins, Regional Staff, other
Admins):

1. **In the app**, as Admin: **User Accounts** → **New User** → fill in
   name, username, **email**, role, and location → **Save User**. This
   creates their profile (no login yet).
2. **In Supabase**: **Authentication** → **Users** → **Add user** → enter
   the *same email* → set a temporary password (or use **Invite** instead
   of **Create user**, if you'd rather they set their own password via an
   emailed link — this is the more standard approach for real employees,
   and avoids you knowing their password).
3. Copy that new user's **User UID**.
4. In **SQL Editor**, link the two:
   ```sql
   update users set auth_user_id = 'PASTE-THE-USER-UID-HERE'
   where email = 'their-email@company.com';
   ```
5. Tell them their login email and (if you set one) temporary password —
   or, if you used **Invite**, they'll get an email directly from
   Supabase with a link to set their own password.

This manual two-step (app + dashboard) is the pragmatic starting point.
Once you have more than a handful of people to onboard, the natural next
step is a small Supabase Edge Function that does steps 2–4 automatically
when an Admin clicks "New User" — worth doing once this is working
end-to-end and you're comfortable with the basics. Happy to help you build
that when you're ready.

---

## Phase 7 — Test everything

- [ ] Admin can log in and see all locations' data.
- [ ] A Regional Staff / Regional Admin account can log in and sees **only**
      their own location's assets, maintenance, comments, and activity log.
- [ ] A Regional Staff / Regional Admin account, if you open the browser
      console and try calling `supabase.from('assets').select('*')`
      directly, only gets rows from their own location back (this is the
      real test that RLS — not just the UI — is enforcing the boundary).
- [ ] Regional Staff cannot delete an asset directly (sees "request
      deletion" instead); Regional Admin and Admin can.
- [ ] Transferring an asset to a different location still works for a
      non-admin (this exercises the "update" RLS policy's edge case).
- [ ] "Reset Password" on a linked user sends a real email and the new
      password works.
- [ ] Signing out and back in works, and refreshing the page mid-session
      keeps you logged in.
- [ ] Someone with no `users` row (sign in with an Auth account you
      haven't linked yet) sees the "Account not set up yet" screen, not a
      crash or blank page.

---

## Phase 8 — Clean-up (once everything above is confirmed working)

Wait until you've had real people successfully log in and use the app for
a few days before doing this — it's not reversible.

```sql
alter table users drop column password_hash;
```

This removes the last leftover trace of the old (insecure) login system
from the database.

---

If everything in Phase 7 checks out, the six critical items are done. The
Recommended-tier items from the architecture review (splitting the file up,
adding tests, etc.) can follow at whatever pace makes sense — none of them
are urgent the way this list was.
