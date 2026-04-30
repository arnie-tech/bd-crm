-- Adds structured Australian addresses for direct mail.
-- Firms hold the primary mailing address; contacts may override per-person.
-- Run this in the Supabase SQL Editor.

ALTER TABLE crm_firms
  ADD COLUMN IF NOT EXISTS address_line1 text,
  ADD COLUMN IF NOT EXISTS address_line2 text,
  ADD COLUMN IF NOT EXISTS suburb        text,
  ADD COLUMN IF NOT EXISTS state         text,
  ADD COLUMN IF NOT EXISTS postcode      text;

-- Existing freeform crm_firms.address is left in place as a legacy field.
-- The app no longer reads or writes it; you can copy values into the
-- structured columns above and drop it later with:
--   ALTER TABLE crm_firms DROP COLUMN address;

ALTER TABLE crm_contacts
  ADD COLUMN IF NOT EXISTS mailing_line1    text,
  ADD COLUMN IF NOT EXISTS mailing_line2    text,
  ADD COLUMN IF NOT EXISTS mailing_suburb   text,
  ADD COLUMN IF NOT EXISTS mailing_state    text,
  ADD COLUMN IF NOT EXISTS mailing_postcode text;
