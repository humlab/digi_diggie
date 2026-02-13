-- Migration script from public schema (old) to digidiggie_tng schema (new)
-- This script migrates data from the old flat structure to the new normalized structure

-- alter schema public rename to digidiggie_tog;
-- alter schema digidiggie_tng rename to digidiggie_tng;

do $$
begin
    raise notice 'Starting migration from old TOG to new TNG digidiggie_tng schema...';
end $$;

-- =============================================================================
-- STEP 1: Migrate simple lookup tables (direct copies)
-- =============================================================================

do $$
begin
    raise notice 'step 1: migrating simple lookup tables...';

    -- communities
    insert into digidiggie_tng.communities (community_id, community_name, parish_id)
        select community_id, community_name, parish_id
        from public.communities
        on conflict (community_id) do nothing;

    -- parishes
    insert into digidiggie_tng.parishes (parish_id, parish)
    select parish_id, parish
    from public.parishes
    on conflict (parish_id) do nothing;

    -- persons
    insert into digidiggie_tng.persons (person_id, given_name, patronymic, surname, birth_year, death_year, community_name)
    select person_id, given_name, patronymic, surname, birth_year, death_year, community_name
    from public.persons
    on conflict (person_id) do nothing;

    -- judgements
    insert into digidiggie_tng.judgements (judgement_id, sanction)
    select judgement_id, sanction
    from public.judgements
    on conflict (judgement_id) do nothing;

    -- land use
    insert into digidiggie_tng.land_use (land_use_id, type)
    select land_use_id, type
    from public.land_use
    on conflict (land_use_id) do nothing;

    -- legal sources
    insert into digidiggie_tng.legal_sources (legal_source_id, legal_source_name)
    select legal_source_id, legal_source_name
    from public.legal_sources
    on conflict (legal_source_id) do nothing;

    -- seasons
    insert into digidiggie_tng.seasons (season_id, season_name)
    select season_id, season_name
    from public.seasons
    on conflict (season_id) do nothing;

    -- sources
    insert into digidiggie_tng.sources (source_id, source_name, source_abbreviation)
    select source_id, source_name, source_abbreviation
    from public.sources
    on conflict (source_id) do nothing;

    -- winners
    insert into digidiggie_tng.winners (winner_id, winner_description)
    select winner_id, winner_description
    from public.winners
    on conflict (winner_id) do nothing;

    -- properties
    insert into digidiggie_tng.properties (property_id, property_name, description)
    select property_id, property_name, description
    from public.properties
    on conflict (property_id) do nothing;

    -- person properties
    insert into digidiggie_tng.person_properties (person_property_id, person_id, property_id, property_value)
    select person_property_id, person_id, property_id, property_value
    from public.person_properties
    on conflict (person_property_id) do nothing;

    -- placenames (if exists)
    insert into digidiggie_tng.placenames (
        fid, objektidentitet, objektversion, objekttypnr, objekttyp, 
        uuid, versiongiltigfran, namn, namntyp, naturrumtyp, language, 
        lan, kommun, socken, geom_point
    )
    select 
        fid, objektidentitet, objektversion, objekttypnr, objekttyp, 
        uuid, versiongiltigfran, namn, namntyp, naturrumtyp, language, 
        lan, kommun, socken, geom_point
    from public.placenames
    where exists (select 1 from information_schema.tables 
                where table_schema = 'public' and table_name = 'placenames')
    on conflict (fid) do nothing;

    raise notice 'step 1 completed: simple lookup tables migrated';
end $$;

-- =============================================================================
-- STEP 2: Create land_right_status lookup table from old data
-- =============================================================================

do $$
begin
    raise notice 'step 2: creating land_right_status lookup table...';

    -- Insert distinct land_rights_status values
    insert into digidiggie_tng.land_right_status (land_rights_status_id, land_rights_status)
    select 
        row_number() over (order by coalesce(land_rights_status, 'unknown')) as land_rights_status_id,
        coalesce(land_rights_status, 'unknown') as land_rights_status
    from (
        select distinct land_rights_status
        from public.entries
        where land_rights_status is not null and land_rights_status != ''
        union
        select 'unknown' -- ensure we have a default value
    ) sub
    on conflict do nothing;

    raise notice 'step 2 completed: land_right_status lookup created';
end $$;

-- =============================================================================
-- STEP 3: Create court_cases from entries
-- =============================================================================

do $$
begin
    raise notice 'step 3: creating court cases...';

    -- create court_cases by grouping entries that belong to the same case
    -- a court case is identified by unique combinations of source_id + reference_number
    insert into digidiggie_tng.court_cases (source_id, reference_number, case_date, source_text)
    select distinct
        coalesce(e.source_id, 1) as source_id,
        e.reference_number,
        case 
            when e.year is not null then make_date(cast(e.year as integer), 1, 1)
            else null
        end as case_date,
        null as source_text -- no direct mapping from old schema
    from public.entries e
    where e.source_id is not null
    on conflict do nothing;

    raise notice 'step 3 completed: court cases created';
end $$;

-- =============================================================================
-- STEP 4: Create entries in new schema
-- =============================================================================

do $$
begin
    raise notice 'step 4: creating entries in new schema...';

    -- create entries linked to court_cases
    insert into digidiggie_tng.entries (entry_id, court_case_id, year, description, season_id, land_use_id, original_placename, placename_id)
    select 
        e.entry_id,
        cc.court_case_id,
        cast(e.year as integer) as year,
        e.description,
        e.season_id,
        e.land_use_id,
        e.original_placename,
        e.placename_id
    from public.entries e
    left join digidiggie_tng.court_cases cc on 
        cc.source_id = coalesce(e.source_id, 1) 
        and (cc.reference_number = e.reference_number 
            or (cc.reference_number is null and e.reference_number is null))
        and (extract(year from cc.case_date) = cast(e.year as integer)
            or (cc.case_date is null and e.year is null))
    where cc.court_case_id is not null
    on conflict (entry_id) do nothing;

    raise notice 'step 4 completed: entries created';
end $$;

-- =============================================================================
-- STEP 5: Create person_entries (person involvement in entries)
-- =============================================================================

do $$
begin
    raise notice 'step 5: creating person_entries...';

    -- create person_entries from old entries
    -- actor_id in old schema becomes the person reference in person_entries
    insert into digidiggie_tng.person_entries (
        entry_id, 
        actor_id, 
        community_id, 
        land_rights_status_id,
        role_id,
        curated_text
    )
    select 
        e.entry_id,
        coalesce(e.actor_id, 0) as actor_id, -- actor_id refers to person_id
        e.community_id,
        coalesce(lrs.land_rights_status_id, 
                (select land_rights_status_id from digidiggie_tng.land_right_status where land_rights_status = 'unknown' limit 1)
        ) as land_rights_status_id,
        null as role_id, -- no direct mapping in old schema
        null as curated_text -- no direct mapping
    from public.entries e
    left join digidiggie_tng.land_right_status lrs on 
        lrs.land_rights_status = coalesce(e.land_rights_status, 'unknown')
    where e.entry_id in (select entry_id from digidiggie_tng.entries)
    and e.actor_id is not null
    on conflict do nothing;

    raise notice 'step 5 completed: person_entries created';
end $$;

-- =============================================================================
-- STEP 6: Create rulings from old entries
-- =============================================================================

do $$
begin
    raise notice 'step 6: creating rulings...';

    -- create rulings for cases that have winner_id, judgement_id, or legal_source_id
    -- note: ruling_type is a new concept, we'll need to populate it separately
    insert into digidiggie_tng.rulings (
        court_case_id,
        ruling_type_id,
        judgement_id,
        legal_source_id
    )
    select distinct
        cc.court_case_id,
        null as ruling_type_id, -- no direct mapping, needs manual population
        e.judgement_id,
        e.legal_source_id
    from public.entries e
    inner join digidiggie_tng.court_cases cc on 
        cc.source_id = coalesce(e.source_id, 1)
        and (cc.reference_number = e.reference_number 
            or (cc.reference_number is null and e.reference_number is null))
        and (extract(year from cc.case_date) = cast(e.year as integer)
            or (cc.case_date is null and e.year is null))
    where (e.judgement_id is not null or e.legal_source_id is not null)
    and cc.court_case_id is not null
    on conflict do nothing;

    raise notice 'step 6 completed: rulings created';
end $$;

-- =============================================================================
-- STEP 7: Update sequences
-- =============================================================================

do $$
begin
    raise notice 'step 7: updating sequences...';

    -- update all sequences to reflect the migrated data
    select setval('digidiggie_tng.communities_community_id_seq', 
        coalesce((select max(community_id) from digidiggie_tng.communities), 1));

    select setval('digidiggie_tng.parishes_parish_id_seq', 
        coalesce((select max(parish_id) from digidiggie_tng.parishes), 1));

    select setval('digidiggie_tng.persons_person_id_seq', 
        coalesce((select max(person_id) from digidiggie_tng.persons), 1));

    select setval('digidiggie_tng.judgements_judgement_id_seq', 
        coalesce((select max(judgement_id) from digidiggie_tng.judgements), 1));

    select setval('digidiggie_tng.land_use_land_use_id_seq', 
        coalesce((select max(land_use_id) from digidiggie_tng.land_use), 1));

    select setval('digidiggie_tng.legal_sources_legal_source_id_seq', 
        coalesce((select max(legal_source_id) from digidiggie_tng.legal_sources), 1));

    select setval('digidiggie_tng.seasons_season_id_seq', 
        coalesce((select max(season_id) from digidiggie_tng.seasons), 1));

    select setval('digidiggie_tng.sources_source_id_seq', 
        coalesce((select max(source_id) from digidiggie_tng.sources), 1));

    select setval('digidiggie_tng.winners_winner_id_seq', 
        coalesce((select max(winner_id) from digidiggie_tng.winners), 1));

    select setval('digidiggie_tng.properties_property_id_seq', 
        coalesce((select max(property_id) from digidiggie_tng.properties), 1));

    select setval('digidiggie_tng.person_properties_person_property_id_seq', 
        coalesce((select max(person_property_id) from digidiggie_tng.person_properties), 1));

    select setval('digidiggie_tng.court_cases_court_case_id_seq', 
        coalesce((select max(court_case_id) from digidiggie_tng.court_cases), 1));

    select setval('digidiggie_tng.entries_entry_id_seq', 
        coalesce((select max(entry_id) from digidiggie_tng.entries), 1));

    select setval('digidiggie_tng.person_entries_person_entry_id_seq', 
        coalesce((select max(person_entry_id) from digidiggie_tng.person_entries), 1));

    select setval('digidiggie_tng.rulings_ruling_id_seq', 
        coalesce((select max(ruling_id) from digidiggie_tng.rulings), 1));

    raise notice 'step 7 completed: sequences updated';
end $$;

-- =============================================================================
-- Summary
-- =============================================================================

do $$
declare
    communities_count integer;
    parishes_count integer;
    persons_count integer;
    entries_count integer;
    court_cases_count integer;
    person_entries_count integer;
begin
    select count(*) into communities_count from digidiggie_tng.communities;
    select count(*) into parishes_count from digidiggie_tng.parishes;
    select count(*) into persons_count from digidiggie_tng.persons;
    select count(*) into entries_count from digidiggie_tng.entries;
    select count(*) into court_cases_count from digidiggie_tng.court_cases;
    select count(*) into person_entries_count from digidiggie_tng.person_entries;
    
    raise notice '========================================';
    raise notice 'Migration Summary:';
    raise notice '========================================';
    raise notice 'Communities migrated: %', communities_count;
    raise notice 'Parishes migrated: %', parishes_count;
    raise notice 'Persons migrated: %', persons_count;
    raise notice 'Court cases created: %', court_cases_count;
    raise notice 'Entries migrated: %', entries_count;
    raise notice 'Person entries created: %', person_entries_count;
    raise notice '========================================';
    raise notice 'Migration completed successfully!';
    raise notice '========================================';
    raise notice '';
    raise notice 'NOTE: The following tables were created but may need manual data entry:';
    raise notice '  - roles (new concept, no source data)';
    raise notice '  - outcome_types (new concept, no source data)';
    raise notice '  - person_outcomes (new concept, no source data)';
    raise notice '  - ruling_type (new concept, no source data)';
    raise notice '  - rulings.ruling_type_id needs to be populated';
end $$;
