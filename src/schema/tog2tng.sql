/***********************************************************************************************************
** Migration script from public schema (old) to digidiggie_tng schema (new)
************************************************************************************************************/

-- alter schema public rename to digidiggie_tog;
-- drop schema if exists digidiggie_tng cascade;

-- create schema digidiggie_tng;

do $$
begin
    raise notice 'Starting migration from old TOG to new TNG digidiggie_tng schema...';
end $$;

/***********************************************************************************************************
** STEP 1: Migrate simple lookup tables (direct copies)
************************************************************************************************************/

do $$
begin
    raise notice 'step 1: migrating simple lookup tables...';

    -- parishes
    insert into digidiggie_tng.parishes (parish_id, parish)
        select parish_id, parish
        from digidiggie_tog.parishes;

    -- communities
    insert into digidiggie_tng.communities (community_id, community_name, parish_id)
        select community_id, community_name, parish_id
        from digidiggie_tog.communities;

    -- persons
    insert into digidiggie_tng.persons (person_id, given_name, patronymic, surname, birth_year, death_year, community_name)
        select person_id, given_name, patronymic, surname, birth_year, death_year, community_name
        from digidiggie_tog.persons;
    
    -- land use
    insert into digidiggie_tng.land_use (land_use_id, description)
    select land_use_id, type as description
    from digidiggie_tog.land_use;

    -- legal sources
    insert into digidiggie_tng.legal_sources (legal_source_id, legal_source_name)
    select legal_source_id, legal_source_name
    from digidiggie_tog.legal_sources;

    -- seasons
    insert into digidiggie_tng.seasons (season_id, season_name)
    select season_id, season_name
    from digidiggie_tog.seasons;

    -- ruling types
    insert into digidiggie_tng.ruling_type (ruling_type_id, ruling_type)
    select judgement_id as ruling_type_id, sanction as ruling_type
    from digidiggie_tog.judgements;

    -- sources
    insert into digidiggie_tng.sources (source_id, source_name, source_abbreviation)
    select source_id, source_name, source_abbreviation
    from digidiggie_tog.sources;

    -- FIXME: #14 Update to correspond with placeaname table in the database (see branch "placenames")
    -- placenames (if exists)
    -- insert into digidiggie_tng.placenames (
    --     fid, objektidentitet, objektversion, objekttypnr, objekttyp, 
    --     uuid, versiongiltigfran, namn, namntyp, naturrumtyp, language, 
    --     lan, kommun, socken, geom_point
    -- )
    -- select 
    --     fid, objektidentitet, objektversion, objekttypnr, objekttyp, 
    --     uuid, versiongiltigfran, namn, namntyp, naturrumtyp, language, 
    --     lan, kommun, socken, geom_point
    -- from digidiggie_tog.placenames
    -- where exists (select 1 from information_schema.tables 
    --             where table_schema = 'digidiggie_tog' and table_name = 'placenames')
    -- on conflict (fid) do nothing;

    raise notice 'step 1 completed: simple lookup tables migrated';
end $$;

/***********************************************************************************************************
** STEP 2: Create land_right_status lookup table from old data
************************************************************************************************************/

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
        from digidiggie_tog.entries
        where land_rights_status is not null and land_rights_status != ''
        union
        select 'unknown' -- ensure we have a default value
    ) sub;

    raise notice 'step 2 completed: land_right_status lookup created';
end $$;

/***********************************************************************************************************
** STEP 3: Create court_cases from entries
************************************************************************************************************/

do $$
begin
    raise notice 'step 3: creating court cases...';

    -- TODO: #17 Add  `district_court_name` to court_cases table and populate it from old schema if possible. No direct mapping in old schema, may require manual population or inference from source_id.
    -- create court_cases by grouping entries that belong to the same case
    -- a court case is identified by unique combinations of source_id + reference_number
    insert into digidiggie_tng.court_cases (source_id, reference_number, case_date, source_text) -- FIXME: #15 create source_text from documets if possible. Concatenate from all curated texts per case in entries if not.
    -- select distinct source_id, reference_number, "year" as case_date, null as source_text
    -- from digidiggie_tog.entries;
    select source_id, reference_number, "year" as case_date, string_agg(distinct description, '; ') -- FIXME: No constraint
    from digidiggie_tog.entries
	group by 1,2,3;

    raise notice 'step 3 completed: court cases created';
end $$;

-- FIXME: #9 Update to handle court_case_id mapping correctly. Fixed?
/***********************************************************************************************************
** STEP 4: Create entries in new schema
************************************************************************************************************/

do $$
begin
    raise notice 'step 4: creating entries in new schema...';

    -- create entries linked to court_cases
    insert into digidiggie_tng.entries (entry_id, court_case_id, "year", curated_text, season_id, land_use_id, original_placename, placename_id)
    select 
        e.entry_id,
        cc.court_case_id,
        cast(e.year as integer) as "year",
        e.description as curated_text,
        e.season_id,
        e.land_use_id,
        e.original_placename,
        e.placename_id
    from digidiggie_tog.entries e
    left join digidiggie_tng.court_cases cc on 
        cc.source_id = e.source_id
      and (cc.reference_number = e.reference_number)
      and (cc.case_date = e.year)
    where cc.court_case_id is not null;

    raise notice 'step 4 completed: entries created';
end $$;

/***********************************************************************************************************
** STEP 5: Create person_entries (person involvement in entries)
************************************************************************************************************/

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
        null as role_id, -- FIXME: #18 Add role_id mapping if possible, otherwise will need manual population. No direct mapping in old schema.
        null as curated_text -- FIXME: #19 Add curated_text if possible, otherwise will need manual population. No direct mapping in old schema.
    from digidiggie_tog.entries e
    left join digidiggie_tng.land_right_status lrs on 
        lrs.land_rights_status = coalesce(e.land_rights_status, 'unknown')
    where e.entry_id in (select entry_id from digidiggie_tng.entries)
    and e.actor_id is not null;

    raise notice 'step 5 completed: person_entries created';
end $$;

/***********************************************************************************************************
** STEP 6: Create rulings from old entries
************************************************************************************************************/

do $$
begin
    raise notice 'step 6: creating rulings...';

    -- create rulings for cases that have winner_id, judgement_id, or legal_source_id
    -- note: ruling_type is a new concept, we'll need to populate it separately
    insert into digidiggie_tng.rulings (
        court_case_id,
        year,
        description,
        ruling_type_id,
        judgement_id, -- FIXME: Not in new schema. This is located in `person_outcomes` in new schema.
        legal_source_id
    )
    select distinct
        cc.court_case_id,
        e.year, -- NOTE: Check,
        null as description, -- no direct mapping, needs manual population
        null as ruling_type_id, -- no direct mapping, needs manual population
        e.judgement_id,
        e.legal_source_id
    from digidiggie_tog.entries e
    inner join digidiggie_tng.court_cases cc on 
        cc.source_id = coalesce(e.source_id, 1)
        and (cc.reference_number = e.reference_number 
            or (cc.reference_number is null and e.reference_number is null))
        and (extract(year from cc.case_date) = cast(e.year as integer)
            or (cc.case_date is null and e.year is null))
    where (e.judgement_id is not null or e.legal_source_id is not null)
    and cc.court_case_id is not null;

    raise notice 'step 6 completed: rulings created';
end $$;

/***********************************************************************************************************
** STEP 7: Update sequences
************************************************************************************************************/

do $$
declare
  r record;
  next_value bigint;
begin
  for r in
    select
      ns.nspname      as schema_name,
      seq.relname     as sequence_name,
      tbl_ns.nspname  as table_schema,
      tbl.relname     as table_name,
      att.attname     as column_name,
      format('%I.%I', ns.nspname, seq.relname) as seq_fqname,
      format('%I.%I', tbl_ns.nspname, tbl.relname) as tbl_fqname
    from pg_class seq
    join pg_namespace ns on ns.oid = seq.relnamespace
    join pg_depend dep on dep.objid = seq.oid
    join pg_class tbl on tbl.oid = dep.refobjid
    join pg_namespace tbl_ns on tbl_ns.oid = tbl.relnamespace
    join pg_attribute att on att.attrelid = tbl.oid and att.attnum = dep.refobjsubid
    where seq.relkind = 'S'
      and dep.deptype IN ('a', 'n')
      and tbl.relkind IN ('r', 'p')
      and ns.nspname = 'digidiggie_tng'
  loop
    execute format( 'SELECT GREATEST(COALESCE(MAX(%1$I), 0) + 1, 1) FROM %2$s', r.column_name, r.tbl_fqname)
        into next_value;
    execute format('SELECT setval(%L::regclass, %s, false)', r.seq_fqname, next_value);
  end loop;
end $$;

/***********************************************************************************************************
** Summary
************************************************************************************************************/

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
    raise notice ' Communities migrated: %', communities_count;
    raise notice ' Parishes migrated: %', parishes_count;
    raise notice ' Persons migrated: %', persons_count;
    raise notice ' Court cases created: %', court_cases_count;
    raise notice ' Entries migrated: %', entries_count;
    raise notice ' Person entries created: %', person_entries_count;
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
