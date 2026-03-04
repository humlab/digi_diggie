/***********************************************************************************************************
** Migration script from public schema (old) to digidiggie_tng schema (new)
************************************************************************************************************/

-- alter schema public rename to digidiggie_tog;
-- drop schema if exists digidiggie_tng cascade;

-- create schema digidiggie_tng;
do $$
begin
    if current_database() <> 'digidiggie' then
         raise exception 'This script must be run in the digidiggie database, current database: %', current_database();
    end if;
end;
$$;

do $$
begin

    raise notice 'Starting migration from old TOG to new TNG digidiggie_tng schema...';

    /***********************************************************************************************************
    ** STEP 1: Migrate simple lookup tables (direct copies)
    ************************************************************************************************************/

    raise notice 'step: migrating lookup tables...';

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
    -- TODO: Remove "Förlikning", "Ekonomisk uppgörelse", "Hänvisning" Då dessa nu finns "Ruling Type"???
    insert into digidiggie_tng.legal_source (legal_source_id, legal_source_name)
        -- select 0, 'Ej angiven'
        -- union
        select legal_source_id, legal_source_name
        from digidiggie_tog.legal_sources
        ;

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
        values (0, '', 'Uppgift ej angiven'),
               (1, 'Dom', 'Dom utfärdad av tinget'),
               (2, 'Förlikning', 'Förlikning mellan parterna'),
               (3, 'Hänvisning', 'Hänvisning till annan domstol eller myndighet'),
               (4, 'Annan', 'Annan typ av domstolsavgörande'),
               (5, 'Okänd', 'Okänd typ av domstolsavgörande'),
               (6, 'Oavgjort', 'Målet slutade oavgjort'),
               (7, 'Ogillas/ingen ändring', 'Målet ogillas eller ingen ändring i tidigare dom'),
               (8, 'Ekonomisk uppgörelse', 'Ekonomisk uppgörelse mellan parterna');

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

    raise notice 'step: migrating placenames...';

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
        "id" as "placename_id",
        "ortnamn" as "placename",
        "n" as "northing",
        "e" as "easting",
        "lopnr"::numeric(10,1)::int as "serial_number",
        "språk_nr" as "language_code",
        "sockenstad_nr" as "parish_code",
        "lan_nr" as "county_code",
        "kommun_nr" as "municipality_code",
        "kombo" as "kombo",
        "sockenstad" as "parish_name",
        st_transform(st_setsrid(st_makepoint(622159, 7286643), 3006), 4326) as geom
    from digidiggie_tog.placenames
      on conflict (placename_id) do nothing;

    /***********************************************************************************************************
    ** Person: Create person_entry from old entries
    ************************************************************************************************************/

    raise notice 'step: migrating persons...';

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

    raise notice 'step: migrating court cases...';

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

    raise notice 'step: migrating court case entries...';

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

    raise notice 'step: migrating person entries...';

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


    /***********************************************************************************************************
    ** STEP     Add person relationships. Currently, the data lacks relationship information.
    ************************************************************************************************************/


    /***********************************************************************************************************
    ** STEP     Create rulings from old entries
    ************************************************************************************************************/

    raise notice 'step: migrating rulings...';
with tog_to_tng as (
		select cc.court_case_id, tng.court_case_entry_id as tng_id, tog.entry_id as tog_id
		from digidiggie_tog.entries tog
		join digidiggie_tng.court_case_entry tng
		  on tog.season_id = tng.season_id
		 and tog.land_use_id = tng.land_use_id
		 and coalesce(tog.placename_id, -1) = coalesce(tng.placename_id, -1)
		join digidiggie_tng.court_case cc
		  on tng.court_case_id = cc.court_case_id
		 and lower(trim(cc.reference_number)) = lower(trim(tog.reference_number))
		 and tog."year"::int = tng.entry_year
    ),
	court_case_ruling_type (source_id, reference_number, case_year, ruling_type) as (values
		(1, '0005', 1649, null),
		(1, '0024', 1673, null),
		(1, '0038a', 1678, 'Hänvisning'),
		(1, '0038b', 1678, 'Okänd'),
		(1, '0045b', 1683, 'Okänd'),
		(1, '0045c', 1682, null),
		(1, '0050b', 1684, 'Dom'),
		(1, '0093', 1691, 'Okänd'),
		(1, '0096', 1691, 'Dom'),
		(1, '0107', 1693, 'Okänd'),
		(1, '0112', 1694, 'Okänd'),
		(1, '0133b', 1697, null),
		(1, '0138b', 1699, 'Okänd'),
		(1, '0153', 1699, 'Okänd'),
		(1, '0158a', 1699, null),
		(1, '0159', 1700, 'Okänd'),
		(1, '0181', 1702, null),
		(1, '0182', 1702, null),
		(1, '0191a', 1702, 'Okänd'),
		(1, '0196', 1703, null),
		(1, '0197', 1703, 'Okänd'),
		(1, '0208', 1704, 'Okänd'),
		(1, '0210', 1705, 'Okänd'),
		(1, '0245', 1705, 'Okänd'),
		(1, '0264', 1706, 'Okänd'),
		(1, '0265', 1706, 'Okänd'),
		(1, '0266', 1706, 'Okänd'),
		(1, '0268', 1706, 'Okänd'),
		(1, '0279', 1707, null),
		(1, '0293', 1709, 'Okänd'),
		(1, '0307', 1710, 'Okänd'),
		(1, '0309', 1710, 'Okänd'),
		(1, '0319', 1711, 'Okänd'),
		(1, '0331', 1712, 'Okänd'),
		(1, '0332', 1712, null),
		(1, '0333', 1690, null),
		(1, '0337', 1703, null),
		(1, '0340', 1711, 'Okänd'),
		(1, '0341', 1713, null),
		(1, '0389', 1716, 'Okänd'),
		(1, '0415', 1720, 'Dom'),
		(1, '0417', 1720, 'Okänd'),
		(1, '0419', 1720, 'Okänd'),
		(1, '0423', 1721, null),
		(1, '0425', 1721, 'Okänd'),
		(1, '0428', 1721, null),
		(1, '0431', 1700, null),
		(1, '0431', 1721, null),
		(1, '0437', 1721, 'Dom'),
		(1, '0441', 1722, null),
		(1, '0445', 1722, null),
		(1, '0447', 1722, 'Förlikning'),
		(1, '0460', 1722, null),
		(1, '0460', 1723, null),
		(1, '0475', 1691, null),
		(1, '0475', 1724, null),
		(1, '0496', 1726, 'Okänd'),
		(1, '0511', 1727, null),
		(1, '0543', 1729, null),
		(1, '0548', 1729, 'Okänd'),
		(1, '0559', 1730, 'Okänd'),
		(1, '0563', 1729, 'Okänd'),
		(1, '0563', 1730, null),
		(1, '0564b', 1730, 'Okänd'),
		(1, '0565', 1730, null),
		(1, '0567', 1730, 'Okänd'),
		(1, '0568', 1729, 'Okänd'),
		(1, '0569', 1730, 'Okänd'),
		(1, '0597', 1731, 'Oavgjort'),
		(1, '0603', 1731, 'Hänvisning'),
		(1, '0605', 1731, 'Okänd'),
		(1, '0607', 1731, null),
		(1, '0659', 1732, 'Okänd'),
		(1, '0661', 1732, 'Okänd'),
		(1, '0662', 1732, 'Okänd'),
		(1, '0663', 1732, 'Okänd'),
		(1, '0664', 1732, 'Okänd'),
		(1, '0665', 1732, 'Okänd'),
		(1, '0671', 1733, null),
		(1, '0672', 1695, 'Okänd'),
		(1, '0672', 1730, 'Okänd'),
		(1, '0672', 1733, 'Okänd'),
		(1, '0674', 1730, 'Dom'),
		(1, '0722', 1733, 'Okänd'),
		(1, '0723', 1733, 'Okänd'),
		(1, '0724', 1733, null),
		(1, '0725', 1733, null),
		(1, '0726', 1733, null),
		(1, '0727', 1733, null),
		(1, '0728', 1733, null),
		(1, '0736', 1733, 'Förlikning'),
		(1, '0750', 1733, null),
		(1, '0775', 1734, 'Ogillas/ingen ändring'),
		(1, '0782', 1734, null),
		(1, '0783', 1734, null),
		(1, '0786', 1734, null),
		(1, '0787', 1734, 'Okänd'),
		(1, '0789', 1718, 'Okänd'),
		(1, '0789', 1734, 'Okänd'),
		(1, '0827', 1735, 'Okänd'),
		(1, '0828', 1735, 'Okänd'),
		(1, '0829', 1735, 'Okänd'),
		(1, '0830', 1735, null),
		(1, '0836', 1735, 'Okänd'),
		(1, '0846', 1736, 'Förlikning'),
		(1, '0849', 1736, 'Ekonomisk uppgörelse'),
		(1, '0851', 1736, 'Okänd'),
		(1, '0852', 1736, 'Okänd'),
		(1, '0854', 1736, null),
		(1, '0855', 1736, null),
		(1, '0860', 1737, 'Förlikning'),
		(1, '0863', 1737, null),
		(1, '0865', 1737, 'Okänd'),
		(1, '0867', 1737, 'Okänd'),
		(1, '0898', 1737, 'Dom'),
		(1, '0902', 1737, 'Ekonomisk uppgörelse'),
		(1, '0910', 1738, 'Dom'),
		(1, '0911', 1738, 'Ekonomisk uppgörelse'),
		(1, '0918', 1739, 'Förlikning'),
		(1, '0923', 1739, 'Hänvisning'),
		(1, '0929', 1739, 'Okänd'),
		(1, '0931', 1739, null),
		(1, '0932', 1739, 'Okänd'),
		(1, '0946', 1740, 'Förlikning'),
		(1, '0960', 1740, null),
		(1, '0961', 1740, 'Okänd'),
		(1, '0969', 1740, 'Förlikning'),
		(1, '0972b', 1739, null),
		(1, '0972b', 1741, 'Okänd'),
		(1, '0977', 1741, 'Förlikning'),
		(1, '0979', 1741, 'Okänd'),
		(1, '0981', 1741, 'Okänd'),
		(1, '0982', 1741, 'Okänd'),
		(1, '0984', 1741, null),
		(1, '0987', 1741, 'Dom'),
		(1, '1000', 1742, 'Hänvisning'),
		(1, '1002', 1742, 'Förlikning'),
		(1, '1007', 1742, 'Okänd'),
		(1, '1009b', 1695, null),
		(1, '1009b', 1742, null),
		(1, '1009e', 1743, 'Okänd'),
		(1, '1010a', 1743, 'Okänd'),
		(1, '1015', 1743, null),
		(1, '1018', 1743, null),
		(1, '1019', 1743, 'Okänd'),
		(1, '1021', 1743, 'Okänd'),
		(1, '1022', 1743, 'Okänd'),
		(1, '1031', 1743, 'Okänd'),
		(1, '1033', 1742, 'Dom'),
		(1, '1045', 1744, null),
		(1, '1046', 1744, null),
		(1, '1054', 1744, null),
		(1, '1055', 1744, 'Oavgjort'),
		(1, '1067', 1745, null),
		(1, '1069', 1745, null),
		(1, '1070', 1745, null),
		(1, '1077e', 1746, 'Förlikning'),
		(1, '1084b', 1748, null),
		(1, '1091', 1748, 'Hänvisning'),
		(1, '1094', 1748, null),
		(1, '1121', 1749, null),
		(1, '1123', 1749, null),
		(1, '1124', 1749, null),
		(1, '1129', 1749, null),
		(1, '1144b', 1750, null),
		(1, '1145', 1750, 'Förlikning'),
		(1, '1149', 1750, 'Förlikning'),
		(1, '1152', 1750, null),
		(1, '1153', 1750, null),
		(1, '1157', 1750, null),
		(1, '1167a', 1751, null),
		(1, '1169', 1751, null),
		(1, '1171', 1751, null),
		(1, '1174b', 1751, null),
		(1, '1174b', 1752, null),
		(1, '1175', 1752, 'Dom'),
		(1, '1177', 1752, 'Dom'),
		(1, '1181', 1752, null),
		(1, '1182', 1752, null),
		(1, '1183', 1752, null),
		(1, '1184', 1752, 'Dom'),
		(1, '1197a', 1753, null),
		(1, '1207c', 1754, null),
		(1, '1209', 1754, 'Förlikning'),
		(1, '1212', 1754, 'Oavgjort'),
		(1, '1214', 1754, 'Dom'),
		(1, '1215', 1754, null),
		(1, '1216', 1754, null),
		(1, '1217', 1754, null),
		(1, '1226b', 1755, null),
		(1, '1227', 1755, 'Okänd'),
		(1, '1232', 1755, null),
		(1, '1238', 1755, 'Dom'),
		(1, '1241b', 1756, null),
		(1, '1253', 1756, null),
		(1, '1254', 1756, null),
		(1, '1255', 1756, null),
		(1, '1256', 1756, null),
		(1, '1257', 1756, null),
		(1, '1258', 1756, null),
		(1, '1266', 1757, 'Dom'),
		(1, '1267', 1757, null),
		(1, '1268', 1757, null),
		(1, '1269', 1757, null),
		(1, '1271', 1757, null),
		(1, '1272', 1757, null),
		(1, '1273', 1757, null),
		(1, '1275', 1757, 'Förlikning'),
		(1, '1278', 1757, 'Okänd'),
		(1, '1287b', 1758, null),
		(1, '1290', 1758, null),
		(1, '1291', 1758, null),
		(1, '1292', 1758, null),
		(1, '1295', 1758, null),
		(1, '1309', 1759, null),
		(1, '1311', 1759, null),
		(1, '1337', 1761, null),
		(1, '1338', 1761, null),
		(1, '1345', 1761, 'Dom'),
		(1, '1346', 1761, 'Ogillas/ingen ändring'),
		(1, '1352', 1762, null),
		(1, '1354', 1762, null),
		(1, '1362c', 1762, null),
		(1, '1362c', 1763, null),
		(1, '1367', 1763, null),
		(1, '1380c', 1763, null),
		(1, '1380c', 1764, null),
		(1, '1384', 1764, null),
		(1, '1397b', 1764, null),
		(1, '1397b', 1765, 'Okänd'),
		(1, '1400', 1765, null),
		(1, '1412', 1766, 'Dom'),
		(1, '1423a', 1767, null),
		(1, '1426', 1767, null),
		(1, '1434b', 1768, null),
		(1, '1446', 1768, null),
		(1, '1451', 1769, null),
		(1, '1451a', 1769, null),
		(1, '1458', 1769, null),
		(1, '1467c', 1770, null),
		(1, '1479', 1770, null),
		(1, '1500a', 1771, null),
		(1, '1501', 1771, null),
		(1, '1519c', 1772, null),
		(1, '1522', 1772, null),
		(1, '1540c', 1773, null),
		(1, '1543', 1773, 'Dom'),
		(1, '1546', 1773, null),
		(1, '1547', 1773, null),
		(1, '1571', 1774, null),
		(1, '1572', 1774, null),
		(1, '1582', 1774, 'Dom'),
		(1, '1591b', 1775, null),
		(1, '1597', 1775, null),
		(1, '1599', 1775, null),
		(1, '1602', 1775, null),
		(1, '1608c', 1776, null),
		(1, '1618', 1776, 'Hänvisning'),
		(1, '1621', 1776, null),
		(1, '1623', 1776, null),
		(1, '1625', 1776, null),
		(1, '1632b', 1777, null),
		(1, '1643', 1777, null),
		(1, '1644', 1777, null),
		(1, '1645', 1777, null),
		(1, '1648', 1777, null),
		(1, '1670', 1778, null),
		(1, '1683c', 1780, null),
		(1, '1727', 1783, 'Hänvisning'),
		(1, '1731', 1783, null),
		(1, '1753', 1784, null),
		(1, '1789', 1785, 'Hänvisning'),
		(1, '1792', 1784, null),
		(1, '1794', 1785, null),
		(1, '1796', 1785, null),
		(1, '1797', 1785, null),
		(1, '1799', 1785, null),
		(1, '1821', 1786, null),
		(1, '1843', 1788, null),
		(1, '1902', 1790, null),
		(1, '1930', 1792, null),
		(1, '1933', 1792, 'Dom'),
		(1, '1955', 1793, 'Ogillas/ingen ändring'),
		(1, '1980', 1773, null),
		(1, '1980', 1794, null),
		(1, '1982', 1794, null),
		(1, '1988', 1794, null),
		(1, '1993', 1794, null),
		(1, '2078', 1797, 'Dom'),
		(1, '2177', 1800, null),
		(1, '2428', 1828, 'Dom'),
		(1, '2435', 1829, 'Dom'),
		(1, '2463', 1835, null),
		(1, '2475', 1836, null),
		(1, '2484', 1837, null),
		(1, '2557', 1844, 'Dom'),
		(1, '546', 1729, null)
    ), legal_sources as (
		select m.court_case_id, max(legal_source_id) as legal_source_id
		from digidiggie_tog.entries e
		join tog_to_tng m
		on m.tog_id = e.entry_id
		group by court_case_id
    )
    insert into digidiggie_tng.ruling ( "court_case_id", "ruling_year", "description", "ruling_type_id", "legal_source_id" )
        select cc."court_case_id",
            cc.case_year as "ruling_year",
            '' as "description",
            coalesce(rt.ruling_type_id, 0) as "ruling_type_id",
            "legal_source_id"
        from court_case_ruling_type ccrt
        join digidiggie_tng.court_case cc using (source_id, reference_number, case_year)
        join legal_sources using (court_case_id)
        left join digidiggie_tng.ruling_type rt using (ruling_type)
        order by cc.court_case_id;


    -- create rulings for cases that have winner_id, judgement_id, or legal_source_id
    -- note: ruling_type is a new concept, we'll need to populate it separately
/*


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
*/

    raise notice 'step: migrating persons'' outcomes...';

end $$;

/***********************************************************************************************************
** STEP 7: Update sequences
************************************************************************************************************/

call digidiggie_tng.sync_sequences()
