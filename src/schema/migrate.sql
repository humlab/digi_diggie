-- Migration script from public schema (old) to digidiggie schema (new)
-- This script migrates data from the old flat structure to the new normalized structure

DO $$
BEGIN
    RAISE NOTICE 'Starting migration from public schema to digidiggie schema...';
END $$;

-- =============================================================================
-- STEP 1: Migrate simple lookup tables (direct copies)
-- =============================================================================

DO $$
BEGIN
    RAISE NOTICE 'Step 1: Migrating simple lookup tables...';
END $$;

-- Communities
INSERT INTO digidiggie.communities (community_id, community_name, parish_id)
SELECT community_id, community_name, parish_id
FROM public.communities
ON CONFLICT (community_id) DO NOTHING;

-- Parishes
INSERT INTO digidiggie.parishes (parish_id, parish)
SELECT parish_id, parish
FROM public.parishes
ON CONFLICT (parish_id) DO NOTHING;

-- Persons
INSERT INTO digidiggie.persons (person_id, given_name, patronymic, surname, birth_year, death_year, community_name)
SELECT person_id, given_name, patronymic, surname, birth_year, death_year, community_name
FROM public.persons
ON CONFLICT (person_id) DO NOTHING;

-- Judgements
INSERT INTO digidiggie.judgements (judgement_id, sanction)
SELECT judgement_id, sanction
FROM public.judgements
ON CONFLICT (judgement_id) DO NOTHING;

-- Land Use
INSERT INTO digidiggie.land_use (land_use_id, type)
SELECT land_use_id, type
FROM public.land_use
ON CONFLICT (land_use_id) DO NOTHING;

-- Legal Sources
INSERT INTO digidiggie.legal_sources (legal_source_id, legal_source_name)
SELECT legal_source_id, legal_source_name
FROM public.legal_sources
ON CONFLICT (legal_source_id) DO NOTHING;

-- Seasons
INSERT INTO digidiggie.seasons (season_id, season_name)
SELECT season_id, season_name
FROM public.seasons
ON CONFLICT (season_id) DO NOTHING;

-- Sources
INSERT INTO digidiggie.sources (source_id, source_name, source_abbreviation)
SELECT source_id, source_name, source_abbreviation
FROM public.sources
ON CONFLICT (source_id) DO NOTHING;

-- Winners
INSERT INTO digidiggie.winners (winner_id, winner_description)
SELECT winner_id, winner_description
FROM public.winners
ON CONFLICT (winner_id) DO NOTHING;

-- Properties
INSERT INTO digidiggie.properties (property_id, property_name, description)
SELECT property_id, property_name, description
FROM public.properties
ON CONFLICT (property_id) DO NOTHING;

-- Person Properties
INSERT INTO digidiggie.person_properties (person_property_id, person_id, property_id, property_value)
SELECT person_property_id, person_id, property_id, property_value
FROM public.person_properties
ON CONFLICT (person_property_id) DO NOTHING;

-- Placenames (if exists)
INSERT INTO digidiggie.placenames (
    fid, objektidentitet, objektversion, objekttypnr, objekttyp, 
    uuid, versiongiltigfran, namn, namntyp, naturrumtyp, language, 
    lan, kommun, socken, geom_point
)
SELECT 
    fid, objektidentitet, objektversion, objekttypnr, objekttyp, 
    uuid, versiongiltigfran, namn, namntyp, naturrumtyp, language, 
    lan, kommun, socken, geom_point
FROM public.placenames
WHERE EXISTS (SELECT 1 FROM information_schema.tables 
              WHERE table_schema = 'public' AND table_name = 'placenames')
ON CONFLICT (fid) DO NOTHING;

DO $$
BEGIN
    RAISE NOTICE 'Step 1 completed: Simple lookup tables migrated';
END $$;

-- =============================================================================
-- STEP 2: Create land_right_status lookup table from old data
-- =============================================================================

DO $$
BEGIN
    RAISE NOTICE 'Step 2: Creating land_right_status lookup table...';
END $$;

-- Insert distinct land_rights_status values
INSERT INTO digidiggie.land_right_status (land_rights_status_id, land_rights_status)
SELECT 
    ROW_NUMBER() OVER (ORDER BY COALESCE(land_rights_status, 'Unknown')) as land_rights_status_id,
    COALESCE(land_rights_status, 'Unknown') as land_rights_status
FROM (
    SELECT DISTINCT land_rights_status
    FROM public.entries
    WHERE land_rights_status IS NOT NULL AND land_rights_status != ''
    UNION
    SELECT 'Unknown' -- ensure we have a default value
) sub
ON CONFLICT DO NOTHING;

DO $$
BEGIN
    RAISE NOTICE 'Step 2 completed: land_right_status lookup created';
END $$;

-- =============================================================================
-- STEP 3: Create court_cases from entries
-- =============================================================================

DO $$
BEGIN
    RAISE NOTICE 'Step 3: Creating court cases...';
END $$;

-- Create court_cases by grouping entries that belong to the same case
-- A court case is identified by unique combinations of source_id + reference_number
INSERT INTO digidiggie.court_cases (source_id, reference_number, case_date, source_text)
SELECT DISTINCT
    COALESCE(e.source_id, 1) as source_id,
    e.reference_number,
    CASE 
        WHEN e.year IS NOT NULL THEN make_date(CAST(e.year AS INTEGER), 1, 1)
        ELSE NULL
    END as case_date,
    NULL as source_text -- no direct mapping from old schema
FROM public.entries e
WHERE e.source_id IS NOT NULL
ON CONFLICT DO NOTHING;

DO $$
BEGIN
    RAISE NOTICE 'Step 3 completed: Court cases created';
END $$;

-- =============================================================================
-- STEP 4: Create entries in new schema
-- =============================================================================

DO $$
BEGIN
    RAISE NOTICE 'Step 4: Creating entries in new schema...';
END $$;

-- Create entries linked to court_cases
INSERT INTO digidiggie.entries (entry_id, court_case_id, year, description, season_id, land_use_id, original_placename, placename_id)
SELECT 
    e.entry_id,
    cc.court_case_id,
    CAST(e.year AS INTEGER) as year,
    e.description,
    e.season_id,
    e.land_use_id,
    e.original_placename,
    e.placename_id
FROM public.entries e
LEFT JOIN digidiggie.court_cases cc ON 
    cc.source_id = COALESCE(e.source_id, 1) 
    AND (cc.reference_number = e.reference_number 
         OR (cc.reference_number IS NULL AND e.reference_number IS NULL))
    AND (EXTRACT(YEAR FROM cc.case_date) = CAST(e.year AS INTEGER)
         OR (cc.case_date IS NULL AND e.year IS NULL))
WHERE cc.court_case_id IS NOT NULL
ON CONFLICT (entry_id) DO NOTHING;

DO $$
BEGIN
    RAISE NOTICE 'Step 4 completed: Entries created';
END $$;

-- =============================================================================
-- STEP 5: Create person_entries (person involvement in entries)
-- =============================================================================

DO $$
BEGIN
    RAISE NOTICE 'Step 5: Creating person_entries...';
END $$;

-- Create person_entries from old entries
-- actor_id in old schema becomes the person reference in person_entries
INSERT INTO digidiggie.person_entries (
    entry_id, 
    actor_id, 
    community_id, 
    land_rights_status_id,
    role_id,
    curated_text
)
SELECT 
    e.entry_id,
    COALESCE(e.actor_id, 0) as actor_id, -- actor_id refers to person_id
    e.community_id,
    COALESCE(lrs.land_rights_status_id, 
             (SELECT land_rights_status_id FROM digidiggie.land_right_status WHERE land_rights_status = 'Unknown' LIMIT 1)
    ) as land_rights_status_id,
    NULL as role_id, -- no direct mapping in old schema
    NULL as curated_text -- no direct mapping
FROM public.entries e
LEFT JOIN digidiggie.land_right_status lrs ON 
    lrs.land_rights_status = COALESCE(e.land_rights_status, 'Unknown')
WHERE e.entry_id IN (SELECT entry_id FROM digidiggie.entries)
  AND e.actor_id IS NOT NULL
ON CONFLICT DO NOTHING;

DO $$
BEGIN
    RAISE NOTICE 'Step 5 completed: person_entries created';
END $$;

-- =============================================================================
-- STEP 6: Create rulings from old entries
-- =============================================================================

DO $$
BEGIN
    RAISE NOTICE 'Step 6: Creating rulings...';
END $$;

-- Create rulings for cases that have winner_id, judgement_id, or legal_source_id
-- Note: ruling_type is a new concept, we'll need to populate it separately
INSERT INTO digidiggie.rulings (
    court_case_id,
    ruling_type_id,
    judgement_id,
    legal_source_id
)
SELECT DISTINCT
    cc.court_case_id,
    NULL as ruling_type_id, -- no direct mapping, needs manual population
    e.judgement_id,
    e.legal_source_id
FROM public.entries e
INNER JOIN digidiggie.court_cases cc ON 
    cc.source_id = COALESCE(e.source_id, 1)
    AND (cc.reference_number = e.reference_number 
         OR (cc.reference_number IS NULL AND e.reference_number IS NULL))
    AND (EXTRACT(YEAR FROM cc.case_date) = CAST(e.year AS INTEGER)
         OR (cc.case_date IS NULL AND e.year IS NULL))
WHERE (e.judgement_id IS NOT NULL OR e.legal_source_id IS NOT NULL)
  AND cc.court_case_id IS NOT NULL
ON CONFLICT DO NOTHING;

DO $$
BEGIN
    RAISE NOTICE 'Step 6 completed: rulings created';
END $$;

-- =============================================================================
-- STEP 7: Update sequences
-- =============================================================================

DO $$
BEGIN
    RAISE NOTICE 'Step 7: Updating sequences...';
END $$;

-- Update all sequences to reflect the migrated data
SELECT setval('digidiggie.communities_community_id_seq', 
    COALESCE((SELECT MAX(community_id) FROM digidiggie.communities), 1));

SELECT setval('digidiggie.parishes_parish_id_seq', 
    COALESCE((SELECT MAX(parish_id) FROM digidiggie.parishes), 1));

SELECT setval('digidiggie.persons_person_id_seq', 
    COALESCE((SELECT MAX(person_id) FROM digidiggie.persons), 1));

SELECT setval('digidiggie.judgements_judgement_id_seq', 
    COALESCE((SELECT MAX(judgement_id) FROM digidiggie.judgements), 1));

SELECT setval('digidiggie.land_use_land_use_id_seq', 
    COALESCE((SELECT MAX(land_use_id) FROM digidiggie.land_use), 1));

SELECT setval('digidiggie.legal_sources_legal_source_id_seq', 
    COALESCE((SELECT MAX(legal_source_id) FROM digidiggie.legal_sources), 1));

SELECT setval('digidiggie.seasons_season_id_seq', 
    COALESCE((SELECT MAX(season_id) FROM digidiggie.seasons), 1));

SELECT setval('digidiggie.sources_source_id_seq', 
    COALESCE((SELECT MAX(source_id) FROM digidiggie.sources), 1));

SELECT setval('digidiggie.winners_winner_id_seq', 
    COALESCE((SELECT MAX(winner_id) FROM digidiggie.winners), 1));

SELECT setval('digidiggie.properties_property_id_seq', 
    COALESCE((SELECT MAX(property_id) FROM digidiggie.properties), 1));

SELECT setval('digidiggie.person_properties_person_property_id_seq', 
    COALESCE((SELECT MAX(person_property_id) FROM digidiggie.person_properties), 1));

SELECT setval('digidiggie.court_cases_court_case_id_seq', 
    COALESCE((SELECT MAX(court_case_id) FROM digidiggie.court_cases), 1));

SELECT setval('digidiggie.entries_entry_id_seq', 
    COALESCE((SELECT MAX(entry_id) FROM digidiggie.entries), 1));

SELECT setval('digidiggie.person_entries_person_entry_id_seq', 
    COALESCE((SELECT MAX(person_entry_id) FROM digidiggie.person_entries), 1));

SELECT setval('digidiggie.rulings_ruling_id_seq', 
    COALESCE((SELECT MAX(ruling_id) FROM digidiggie.rulings), 1));

DO $$
BEGIN
    RAISE NOTICE 'Step 7 completed: Sequences updated';
END $$;

-- =============================================================================
-- Summary
-- =============================================================================

DO $$
DECLARE
    communities_count INTEGER;
    parishes_count INTEGER;
    persons_count INTEGER;
    entries_count INTEGER;
    court_cases_count INTEGER;
    person_entries_count INTEGER;
BEGIN
    SELECT COUNT(*) INTO communities_count FROM digidiggie.communities;
    SELECT COUNT(*) INTO parishes_count FROM digidiggie.parishes;
    SELECT COUNT(*) INTO persons_count FROM digidiggie.persons;
    SELECT COUNT(*) INTO entries_count FROM digidiggie.entries;
    SELECT COUNT(*) INTO court_cases_count FROM digidiggie.court_cases;
    SELECT COUNT(*) INTO person_entries_count FROM digidiggie.person_entries;
    
    RAISE NOTICE '========================================';
    RAISE NOTICE 'Migration Summary:';
    RAISE NOTICE '========================================';
    RAISE NOTICE 'Communities migrated: %', communities_count;
    RAISE NOTICE 'Parishes migrated: %', parishes_count;
    RAISE NOTICE 'Persons migrated: %', persons_count;
    RAISE NOTICE 'Court cases created: %', court_cases_count;
    RAISE NOTICE 'Entries migrated: %', entries_count;
    RAISE NOTICE 'Person entries created: %', person_entries_count;
    RAISE NOTICE '========================================';
    RAISE NOTICE 'Migration completed successfully!';
    RAISE NOTICE '========================================';
    RAISE NOTICE '';
    RAISE NOTICE 'NOTE: The following tables were created but may need manual data entry:';
    RAISE NOTICE '  - roles (new concept, no source data)';
    RAISE NOTICE '  - outcome_types (new concept, no source data)';
    RAISE NOTICE '  - person_outcomes (new concept, no source data)';
    RAISE NOTICE '  - ruling_type (new concept, no source data)';
    RAISE NOTICE '  - rulings.ruling_type_id needs to be populated';
END $$;
