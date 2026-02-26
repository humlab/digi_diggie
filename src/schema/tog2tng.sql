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
        -- select 0, 'Ej angiven'
        -- union
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

    -- Relationship Type
    insert into digidiggie_tng.relationship_type ("relationship_type_id", "relationship_type_name", "description")
        values (1, 'Far', 'Person är far till person'),
               (2, 'Mor', 'Person är mor till person'),
               (3, 'Syskon', 'Person är syskon till person'),
               (4, 'Make/Maka', 'Person är make/maka till person'),
               (5, 'Annan', 'Annan typ av relation mellan personer'),
               (0, 'Okänd', 'Okänd typ av relation mellan personer');

    /***********************************************************************************************************
    ** STEP     Populate placename
    ** FIXME    Can this table be reduced e.g. only placenames in northern Sweden?
    ************************************************************************************************************/

    -- placenames
    insert into digidiggie_tng.placename (
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

    insert into digidiggie_tng.court_case (source_id, reference_number, case_year, source_text) -- FIXME: #15 create source_text from documets if possible. Concatenate from all curated texts per case in entries if not.
        select source_id, reference_number, cast("year" as integer) as case_year, string_agg(distinct description, '; ') -- FIXME: No constraint
        from digidiggie_tog.entries
        group by 1, 2, 3;

    /***********************************************************************************************************
    ** STEP     Create court case entries
    ** FIXME    Can entry_year vary over season/land_use/placename???
    ** NOTE     We need to aggregate description and original_placename to avoid duplicates when
    **          multiple entries link to the same case/season/land_use/placename combination.
    ************************************************************************************************************/

    insert into digidiggie_tng.court_case_entry (court_case_id, entry_year, season_id, land_use_id, placename_id, curated_text, original_placename) --original_placename)
        select
            -- Court Case:
            cc.court_case_id                                as court_case_id,
            cast(e.year as integer)                         as entry_year,
            -- Entry keys
            e.season_id                                     as season_id,
            e.land_use_id                                   as land_use_id,
            e.placename_id                                  as placename_id,
            -- extra information about the entry: (must be aggregated to avoid duplicates)
            string_agg(distinct e.description, ';')         as curated_text,
            string_agg(distinct e.original_placename, ';')  as original_placename
        from digidiggie_tng.court_case cc
        join digidiggie_tog.entries e
          on e.source_id = cc.source_id
         and e.reference_number = cc.reference_number
         and cast(e.year as integer) = cc.case_year
        group by cc.court_case_id, e.year, e.season_id, e.land_use_id, e.placename_id
        ;

    
    /***********************************************************************************************************
    ** Person Case Entries: Create person_entry from old entries
    ** FIXME: Many NULL values in role_id
    ** FIXME: Perhaps role_id should be not null, and set to "not specified" if NULL
    ************************************************************************************************************/

    insert into digidiggie_tng.person_entry (
        court_case_entry_id, 
        person_id, 
        community_id, 
        land_rights_status_id,
        role_id,
        curated_text
    )
        with winner_to_role (winner_id, role_id) as (values 
            (1, 7), --	Nybyggare
            (2, 8), --	Bonde
            (3, 6), --	Same
            (4, 9) --	Renskötare
        ), tog_to_tng as (
            select tog.entry_id as tog_id, tng.court_case_entry_id as tng_id
            from digidiggie_tog.entries tog
            join digidiggie_tng.court_case_entry tng
              on tog.season_id = tng.season_id
             and tog.land_use_id = tng.land_use_id
             and coalesce(tog.placename_id, -1) = coalesce(tng.placename_id, -1)
            join digidiggie_tng.court_case cc
              on tng.court_case_id = cc.court_case_id
             and lower(trim(cc.reference_number)) = lower(trim(tog.reference_number))
             and tog."year"::int = tng.entry_year
        )
        select 
            m.tng_id as court_case_entry_id,
            coalesce(tog.actor_id, 0) as person_id,
            tog.community_id as community_id,
            lrs.land_rights_status_id as land_rights_status_id,
            r.role_id as role_id,
            tog.description as curated_text
        from digidiggie_tog.entries tog
		join tog_to_tng m
		  on m.tog_id = tog.entry_id
        join digidiggie_tng.land_rights_status lrs
          on upper(lrs.land_rights_status) = upper(trim(tog.land_rights_status))
        left join winner_to_role r
          on r.winner_id = tog.winner_id;


    -- Winners svarar på frågan vem som vann????
    select *
    from digidiggie_tog.entries
    where winner_id in (5,6,7)

    select *
    from digidiggie_tng.ruling r
    insert into digidiggie_tng.ruling (
        "court_case_id",
        "ruling_year",
        "description",
        "ruling_type_id",
        "legal_source_id"
    ) 
        select cc.court_case_id, cc.case_year, null, null --, e.legal_source_id
        from digidiggie_tng.court_case cc

    /***********************************************************************************************************
    ** STEP     Add person relationships. Currently, the data lacks relationship information.
    ************************************************************************************************************/


    /***********************************************************************************************************
    ** STEP     Create rulings from old entries
    ************************************************************************************************************/

    -- create rulings for cases that have winner_id, judgement_id, or legal_source_id
    -- note: ruling_type is a new concept, we'll need to populate it separately
    insert into digidiggie_tng.rulings (
        "court_case_id",
        "ruling_year",
        "description",
        "ruling_type_id",
        "legal_source_id"
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

call digidiggie_tng.sync_sequences()
