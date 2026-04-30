-- Mailing lists for direct mail. Static snapshots — members are added
-- explicitly and don't change unless you edit the list.
-- Run this in the Supabase SQL Editor.

CREATE TABLE IF NOT EXISTS crm_lists (
  id          uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name        text NOT NULL,
  description text,
  created_at  timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS crm_list_members (
  list_id    uuid NOT NULL REFERENCES crm_lists(id)    ON DELETE CASCADE,
  contact_id uuid NOT NULL REFERENCES crm_contacts(id) ON DELETE CASCADE,
  added_at   timestamptz NOT NULL DEFAULT now(),
  PRIMARY KEY (list_id, contact_id)
);

CREATE INDEX IF NOT EXISTS crm_list_members_contact_idx
  ON crm_list_members (contact_id);

-- Match the RLS posture chosen in 0002_rls.sql (currently: disabled).
ALTER TABLE crm_lists         DISABLE ROW LEVEL SECURITY;
ALTER TABLE crm_list_members  DISABLE ROW LEVEL SECURITY;
