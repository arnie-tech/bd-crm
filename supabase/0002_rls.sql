-- Run ONE of the two options below in the Supabase SQL Editor.
--
-- Background: by default Supabase enables Row Level Security on every
-- table but ships no policies, so every read and write through the anon
-- key gets rejected. The error you see in the app is:
--   "new row violates row-level security policy for table ..."
--
-- Pick the option that matches how you intend to secure this app.

-- ============================================================
-- OPTION A — Disable RLS (simplest, no auth)
-- ============================================================
-- Use this while the app has no login screen. Your Supabase anon key
-- becomes the only thing guarding the data. Treat it like a password:
-- don't post it publicly, rotate it via the Supabase dashboard if it
-- leaks. Easy to switch to Option B later once you add Supabase Auth.

ALTER TABLE crm_contacts     DISABLE ROW LEVEL SECURITY;
ALTER TABLE crm_firms        DISABLE ROW LEVEL SECURITY;
ALTER TABLE crm_activities   DISABLE ROW LEVEL SECURITY;
ALTER TABLE crm_tags         DISABLE ROW LEVEL SECURITY;
ALTER TABLE crm_contact_tags DISABLE ROW LEVEL SECURITY;


-- ============================================================
-- OPTION B — Keep RLS, allow anon full access (functionally same)
-- ============================================================
-- Same effective security as Option A, but RLS stays on so you can
-- layer per-user policies later without an extra ALTER TABLE step.
-- If you take this option, comment out Option A above and uncomment
-- the block below.
--
-- DO $$
-- DECLARE t text;
-- BEGIN
--   FOR t IN SELECT unnest(ARRAY['crm_contacts','crm_firms','crm_activities','crm_tags','crm_contact_tags']) LOOP
--     EXECUTE format('DROP POLICY IF EXISTS "allow anon all" ON %I', t);
--     EXECUTE format(
--       'CREATE POLICY "allow anon all" ON %I FOR ALL TO anon USING (true) WITH CHECK (true)', t
--     );
--   END LOOP;
-- END $$;
