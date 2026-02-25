/***********************************************************************************************************
** Migration script from public schema (old) to digidiggie_tng schema (new)
************************************************************************************************************/

-- alter schema public rename to digidiggie_tog;
-- drop schema if exists digidiggie_tng cascade;

-- create schema digidiggie_tng;


do $$
begin

    raise notice 'Starting migration from old TOG to new TNG digidiggie_tng schema...';

    /***********************************************************************************************************
    ** STEP 1: Migrate simple lookup tables (direct copies)
    ************************************************************************************************************/

    raise notice 'step 1: migrating simple lookup tables...';

    -- parishes
    insert into digidiggie_tng.parish (parish_id, parish)
        select parish_id, parish
        from digidiggie_tog.parishes;

    -- communities
    insert into digidiggie_tng.community (community_id, community_name, parish_id)
        select community_id, community_name, parish_id
        from digidiggie_tog.communities;
    
    -- land use
    insert into digidiggie_tng.land_use (land_use_id, description)
        select land_use_id, type as description
        from digidiggie_tog.land_use;

    -- legal sources
    insert into digidiggie_tng.legal_source (legal_source_id, legal_source_name)
        select legal_source_id, legal_source_name
        from digidiggie_tog.legal_sources;

    -- seasons
    insert into digidiggie_tng.season (season_id, season_name)
        select season_id, season_name
        from digidiggie_tog.seasons;

    -- sources
    insert into digidiggie_tng.source (source_id, source_name, source_abbreviation)
        select source_id, source_name, source_abbreviation
        from digidiggie_tog.sources;

    -- land right status
    insert into digidiggie_tng.land_rights_status (land_rights_status_id, land_rights_status, description)
        values (1, 'Nja', 'Land rights status uncertain'),
               (2, 'Ja', 'Owned land'),
               (3, 'Nej', 'Not owned land');

    -- Ruling types
    insert into digidiggie_tng.ruling_type ("ruling_type_id", "ruling_type", "description")
        values (1, 'Dom', 'Dom utfärdad av tinget'),
               (2, 'Förlikning', 'Förlikning mellan parterna'),
               (3, 'Hänvisning', 'Hänvisning till annan domstol eller myndighet'),
               (4, 'Annan', 'Annan typ av domstolsavgörande'),
               (5, 'Okänd', 'Okänd typ av domstolsavgörande');

    -- Role Type
    insert into digidiggie_tng.role_type ("role_type_id", "role_type_name", "description")
        values (1, 'Social', 'Social roll, t.ex. bonde, kyrkoherde, fogde'),
               (2, 'Juridisk', 'Juridisk roll, t.ex. klagande, svarande, vittne');

    -- Role
    insert into digidiggie_tng.role ("role_id", "role_type_id", "role_name", "description")
        values (1, 2, 'Klagande', 'Part som klagar på en annan part'),
               (2, 2, 'Svarande', 'Part som svarar på en klagan'),
               (3, 2, 'Vittne', 'Person som vittnar i en rättegång'),
               (4, 2, 'Annan', 'Annan roll i en rättegång'),
               (5, 2, 'Okänd', 'Okänd roll i en rättegång'),
               (6, 1, 'Same', 'Person som är same'),
               (7, 1, 'Nybyggare', 'Person som är nybyggare'),
               (8, 1, 'Bonde', 'Person som är bonde'),
               (9, 1, 'Renskötare', 'Person som är renskötare');

    -- Outcome Type
    insert into digidiggie_tng.outcome_type ("outcome_type_id", "outcome_type_name", "description")
        values
            (0, 'Okänd', 'Okänd utgång av målet'),
            (1, 'Skadestånd', 'Personen förlorade målet'),  -- 'Svarande'
            (2, 'Vite', 'Personen fick vite'),              -- 'Svarande'
            (3, 'Böter', 'Personen fick böter'),            -- 'Svarande'
            (4, 'Friad', 'Personen blev friad'),            -- 'Svarande'
            (5, 'Döden', 'Personen dömdes till döden'),     -- 'Svarande'
            (6, 'Fängelse', 'Personen fick fängelse'),      -- 'Svarande'
            (7, 'Förmaning', 'Personen fick förmaning'),    -- 'Svarande'
            (8, 'Annan', 'Annan utgång av målet'),
            (9, 'Vinnare', 'Personen vann målet'),          -- 'Klagande'
            (10, 'Oavgjort', 'Målet slutade oavgjort');

    /***********************************************************************************************************
    ** STEP    Populate placenames (if exists))
    ************************************************************************************************************/

    -- FIXME: #14 Update to correspond with placeaname table in the database (see branch "placenames")
    -- placenames (if exists)
    insert into digidiggie_tng.placenames (
        "placename_id",
        "placename",
        "northing",
        "easting",
        "serial_number",
        "name_type_code",
        "language_code",
        "parish_code",
        "county_code",
        "municipality_code",
        "combined_placename",
        "parish_name"
    )
    select 
        "fid" as "placename_id",
        "ortnamn" as "placename",
        "nkoordinat" as "northing",
        "ekoordinat" as "easting",
        "lopnummer"::numeric(10,1)::int as "serial_number",
        "detaljtyp" as "name_type_code",
        "sprak" as "language_code",
        "sockenstadkod" as "parish_code",
        "lanskod" as "county_code",
        "kommunkod" as "municipality_code",
        -- "kombo" as "combined_placename",
        "sockenstadnamn" as "parish_name",
        st_transform(st_setsrid(st_makepoint(622159, 7286643), 3006), 4326) as geom
    from digidiggie_tog.placenames
    on conflict (placename_id) do nothing;

    /***********************************************************************************************************
    ** Person: Create person_entry from old entries
    ************************************************************************************************************/

    insert into digidiggie_tng.person (person_id, given_name, patronymic, surname, birth_year, death_year, community_name)
        select person_id, given_name, patronymic, surname, birth_year, death_year, community_name
        from digidiggie_tog.persons;
 
    insert into digidiggie_tng.person (person_id, given_name, patronymic, surname, birth_year, death_year, community_name)
        values (0, 'Ej namngiven', '', '', null, null, null);
        
    /***********************************************************************************************************
    ** STEP     Create court_cases from entries
    **          A Court Case SHOULD be uniquely identified by source_id + reference_number
    ** FIX      Added "case_year" to uniquely identify cases
    ** FIX      Case "description" is now aggregated from all entries linked to the same case, separated by "; "
    **          It should however be noted that the case description is supposed to be true to the source.
    ** TODD     Where does "district_courrt_name" come from? Is it in the source data? If not, we can leave it null for now and populate it later if needed.
    ************************************************************************************************************/

    insert into digidiggie_tng.court_cases (source_id, reference_number, case_year, source_text) -- FIXME: #15 create source_text from documets if possible. Concatenate from all curated texts per case in entries if not.
        select source_id, reference_number, "year" as case_year, string_agg(distinct description, '; ') -- FIXME: No constraint
        from digidiggie_tog.entries
        group by 1, 2, 3;

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
      and (cc.case_year = e.year)
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
