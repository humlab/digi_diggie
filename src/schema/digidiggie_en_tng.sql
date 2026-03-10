-- Enable PostGIS extension for geometry types in public schema
do $$
begin
    if current_database() <> 'digidiggie' then
         raise exception 'This script must be run in the digidiggie database, current database: %', current_database();
    end if;
end;
$$;

do $$
begin
    raise notice 'Ensuring PostGIS extension is installed in public schema...';
    if not exists (select 1 from pg_extension where extname = 'postgis') then
        create extension postgis schema public;
    end if;
    drop schema if exists digidiggie_tng cascade;
end;
$$;

do $$
begin
    create schema digidiggie_tng;
end;
$$;

set search_path to digidiggie_tng;
set role gudrun;

create table community (
    "community_id" serial not null primary key,
    "community_name" text not null default ''::text,
    "parish_id" integer
);

create table court_case (
    "court_case_id" serial primary key,
    "source_id" integer not null,
    "reference_number" varchar(16),
    "district_court_name" text,
    "case_year" integer,
    "source_text" text,
    constraint "court_cases_source_id_key" unique ("source_id", "reference_number", "case_year")
);

create table court_case_entry (
    "court_case_entry_id" serial primary key,
    "court_case_id" integer not null,
    "entry_year" integer,
    "curated_text" text,
    "original_placename" text,
    "season_id" integer,
    "land_use_id" integer,
    "placename_id" integer
);

create table person (
    "person_id" serial primary key,
    "given_name" text,
    "patronymic" text,
    "surname" text,
    "birth_year" integer,
    "death_year" integer,
    "community_name" text,
    "note" text null,
    "full_name" text generated always as (
        ((((coalesce("given_name", ''::text) || ' '::text) ||
            coalesce("patronymic", ''::text)) || ' '::text) ||
            coalesce("surname", ''::text))
    ) stored
);

create table land_rights_status (
    "land_rights_status_id" serial primary key,
    "land_rights_status" varchar(255) not null,
    "description" text not null
);

create table legal_source (
    "legal_source_id" serial primary key,
    "legal_source_name" text default ''::text not null
);

create table land_use (
    "land_use_id" serial primary key,
    "description" text default ''::text not null
);

create table parish (
    "parish_id" serial primary key,
    "parish" text not null default ''::text
);

create table outcome_type (
    "outcome_type_id" serial primary key,
    "outcome_type_name" text not null,
    "description" text,
    constraint "outcome_types_outcome_type_name_key" unique ("outcome_type_name")
);

create table person_outcome (
    "person_outcome_id" serial primary key,
    "ruling_id" integer not null,
    "person_id" integer not null,
    "outcome_type_id" integer not null,
    "description" text
);

create table  placename (
    "placename_id" serial primary key,  --> id
    "placename" text,                   --> ortnamn
    "northing" integer,                 --> n SWEREF 99 TM (EPSG:3006)
    "easting" integer,                  --> e SWEREF 99 TM (EPSG:3006)
    "serial_number" text,               --> lopnr
    "name_type_code" text,              --> namntyp_nr
    "language_code" text,               --> språk_nr
    "parish_code" text,                 --> sockenstad_nr
    "county_code" text,                 --> lan_nr
    "municipality_code" text,           --> kommun_nr
    "combined_placename" text,          --> kombo TODO: Check
    "parish_name" text,
    "geom" public.geometry(Point, 4326) --> WGS84 (EPSG:4326)
);

create table person_entry (
    "person_entry_id" serial primary key,
    "court_case_entry_id" integer not null,
    "person_id" integer not null,
    "community_id" integer,
    "land_rights_status_id" integer not null,
    "role_id" integer, 
    -- "judicial_role_id" integer, 
    -- "social_role_id" integer, 
    "curated_text" text
);

create table role_type (
    "role_type_id" serial primary key,
    "role_type_name" text not null,
    "description" text not null,
    constraint "role_types_role_type_name_key" unique ("role_type_name")
);

create table role (
    "role_id" serial primary key,
    "role_name" text not null,
    "role_type_id" integer not null,
    "description" text not null,
    constraint "roles_role_name_key" unique ("role_name")
);

create table season (
    "season_id" serial primary key,
    "season_name" text not null default ''::text,
    constraint "seasons_season_name_key" unique ("season_name")
);

create table ruling_type (
    "ruling_type_id" serial primary key,
    "ruling_type" varchar(255) not null,
    "description" text not null,
    constraint "ruling_types_ruling_type_key" unique ("ruling_type")
);

create table ruling (
    "ruling_id" serial primary key,
    "court_case_id" integer not null,
    "ruling_year" integer, 
    "description" text,
    "ruling_type_id" integer not null,
    "legal_source_id" integer,
    constraint "rulings_court_case_id_key" unique ("court_case_id")
);

create table source (
    "source_id" serial primary key,
    "source_name" text not null default ''::text,
    "source_abbreviation" varchar(255)
);

create table person_relationship (
    "person_relationship_id" serial primary key,
    "person_1_id" integer not null,
    "person_2_id" integer not null,
    "relationship_type_id" integer not null,
    "description" text
);

create table relationship_type (
    "relationship_type_id" serial primary key,
    "relationship_type_name" text not null,
    "description" text not null,
    constraint "relationship_types_relationship_type_name_key" unique ("relationship_type_name")
);

/***********************************************************************************************************
** Foreign key constraints
************************************************************************************************************/

alter table "community" add constraint "community_parish_id_fkey" foreign key ("parish_id") references "parish" ("parish_id");
alter table "court_case" add constraint "court_case_source_id_fkey" foreign key ("source_id") references "source" ("source_id");
alter table "court_case_entry" add constraint "entry_court_case_id_fkey" foreign key ("court_case_id") references "court_case" ("court_case_id");
alter table "court_case_entry" add constraint "entry_land_use_id_fkey" foreign key ("land_use_id") references "land_use" ("land_use_id");
alter table "court_case_entry" add constraint "entry_season_id_fkey" foreign key ("season_id") references "season" ("season_id");
alter table "court_case_entry" add constraint "entry_placename_id_fkey" foreign key ("placename_id") references "placename" ("placename_id");
alter table "person_entry" add constraint "person_entry_person_id_fkey" foreign key ("person_id") references "person" ("person_id");
alter table "person_entry" add constraint "person_entry_community_id_fkey" foreign key ("community_id") references "community" ("community_id");
alter table "person_entry" add constraint "person_entry_entry_id_fkey" foreign key ("court_case_entry_id") references "court_case_entry" ("court_case_entry_id");
alter table "person_entry" add constraint "person_entry_role_id_fkey" foreign key ("role_id") references "role" ("role_id");
alter table "person_entry" add constraint "fk_person_entry_land_rights_status_1" foreign key ("land_rights_status_id") references "land_rights_status" ("land_rights_status_id");
alter table "role" add constraint "role_role_type_id_fkey" foreign key ("role_type_id") references "role_type" ("role_type_id");
alter table "person_outcome" add constraint "person_outcomes_outcome_type_id_fkey" foreign key ("outcome_type_id") references "outcome_type" ("outcome_type_id") on delete no action on update no action;
alter table "person_outcome" add constraint "person_outcomes_person_id_fkey" foreign key ("person_id") references "person" ("person_id") on delete no action on update no action;
alter table "person_outcome" add constraint "person_outcomes_ruling_id_fkey" foreign key ("ruling_id") references "ruling" ("ruling_id") on delete cascade on update no action;
alter table "ruling" add constraint "rulings_court_case_id_fkey" foreign key ("court_case_id") references "court_case" ("court_case_id") on delete cascade on update no action;
alter table "ruling" add constraint "rulings_legal_source_id_fkey" foreign key ("legal_source_id") references "legal_source" ("legal_source_id") on delete no action on update no action;
alter table "ruling" add constraint "fk_rulings_ruling_type_1" foreign key ("ruling_type_id") references "ruling_type" ("ruling_type_id");
alter table "person_relationship" add constraint "person_relationship_person_1_id_fkey" foreign key ("person_1_id") references "person" ("person_id");
alter table "person_relationship" add constraint "person_relationship_person_2_id_fkey" foreign key ("person_2_id") references "person" ("person_id");
alter table "person_relationship" add constraint "person_relationship_relationship_type_id_fkey" foreign key ("relationship_type_id") references "relationship_type" ("relationship_type_id");

/***********************************************************************************************************
** Indexes
************************************************************************************************************/

create index communities_parish_id_idx on "community" ("parish_id");
create index court_cases_case_year_idx on "court_case" ("case_year");
create index court_cases_reference_number_idx on "court_case" ("reference_number");
create index court_cases_source_id_idx on "court_case" ("source_id");
create index entries_court_case_id_idx on "court_case_entry" ("court_case_id");
create index entries_land_use_id_idx on "court_case_entry" ("land_use_id");
create index entries_placename_id_idx on "court_case_entry" ("placename_id");
create index rulings_legal_source_id_idx on "ruling" ("legal_source_id");
create index entries_season_id_idx on "court_case_entry" ("season_id");
create index parishes_parish_idx on "parish" ("parish");
create index person_entries_person_id_idx on "person_entry" ("person_id");
create index person_entries_community_id_idx on "person_entry" ("community_id");
create index person_entries_court_case_entry_id_idx on "person_entry" ("court_case_entry_id");
create index person_outcomes_outcome_type_id_idx on "person_outcome" ("outcome_type_id");
create index person_outcomes_person_id_idx on "person_outcome" ("person_id");
create index person_outcomes_ruling_id_idx on "person_outcome" ("ruling_id");
create index seasons_season_name_idx on "season" ("season_name");


/***********************************************************************************************************
** Database comments
************************************************************************************************************/

comment on schema digidiggie_tng is 'DigiDiggie The Next Generation schema - normalized structure for Swedish court records';

-- Community table
comment on table community is 'Social or administrative grouping (e.g., parish or community) to which a person belongs at the time of the court case';
comment on column community.community_id is 'Primary key for the community table';
comment on column community.community_name is 'The name of the community (village)';
comment on column community.parish_id is 'Foreign key linking to the parish this community belongs to';

-- Court case table
comment on table court_case is 'Single court proceeding recorded in a historical source, identified by date, court, and source text. A court case results in exactly one ruling in this model';
comment on column court_case.court_case_id is 'Primary key for the court case table';
comment on column court_case.source_id is 'Foreign key to the source document (e.g., court record collection)';
comment on column court_case.reference_number is 'Reference number within a specific source collection, e.g., K.B. Wiklund''s transcripts';
comment on column court_case.district_court_name is 'Name of the district court (tingslag) where the case was heard';
comment on column court_case.case_year is 'The year the case was heard at court';
comment on column court_case.source_text is 'Aggregated description/text from the source document for this case';

-- Court case entry table
comment on table court_case_entry is 'Discrete unit of information extracted from a court case, typically describing a specific land-use situation, event, or claim';
comment on column court_case_entry.court_case_entry_id is 'Primary key for the court case entry table';
comment on column court_case_entry.court_case_id is 'Foreign key linking to the parent court case';
comment on column court_case_entry.entry_year is 'The year the event occurred, or if unknown, the year the matter was heard at court';
comment on column court_case_entry.curated_text is 'Curated description of the event in free text';
comment on column court_case_entry.original_placename is 'The place''s name as written in the original source document';
comment on column court_case_entry.season_id is 'Foreign key to the season when the disputed resource was used';
comment on column court_case_entry.land_use_id is 'Foreign key to the resource or land use type involved in the dispute';
comment on column court_case_entry.placename_id is 'Foreign key to a standardized placename from the Swedish National Survey (enables GIS connection)';

-- Person table
comment on table person is 'Historical individual identified in the sources, with personal attributes where known (name, patronymic, birth/death year, notes)';
comment on column person.person_id is 'Primary key for the person table';
comment on column person.given_name is 'The person''s given name(s) (first name)';
comment on column person.patronymic is 'The person''s patronymic name (e.g., ''Andersson'', ''Jonsdotter'')';
comment on column person.surname is 'The person''s surname, byname, or family name';
comment on column person.birth_year is 'The year of birth';
comment on column person.death_year is 'The year of death';
comment on column person.community_name is 'The name of the village where the person primarily resided';
comment on column person.note is 'Additional notes about the person';
comment on column person.full_name is 'The person''s full name, computed from given name, patronymic, and surname';

-- Person entry table
comment on table person_entry is 'Contextualized appearance of a person within a specific court case entry. Captures the person''s role, land rights status, and how they are described in the source';
comment on column person_entry.person_entry_id is 'Primary key for the person entry table';
comment on column person_entry.court_case_entry_id is 'Foreign key linking to the court case entry';
comment on column person_entry.person_id is 'Foreign key linking to the person involved';
comment on column person_entry.community_id is 'Foreign key to the community where the person resided at the time';
comment on column person_entry.land_rights_status_id is 'Foreign key indicating if the person had land rights';
comment on column person_entry.role_id is 'Foreign key to the person''s role in the case (e.g., plaintiff, defendant, witness)';
comment on column person_entry.curated_text is 'Additional curated text describing the person''s involvement';

-- Ruling table
comment on table ruling is 'Judicial decision resulting from a court case. Records the year, description, ruling type (resolved or referred), and may cite a legal source';
comment on column ruling.ruling_id is 'Primary key for the ruling table';
comment on column ruling.court_case_id is 'Foreign key linking to the court case (one-to-one relationship)';
comment on column ruling.ruling_year is 'The year the ruling was issued';
comment on column ruling.description is 'Description of the ruling in free text';
comment on column ruling.ruling_type_id is 'Foreign key to the type of ruling (e.g., judgement, settlement, referral)';
comment on column ruling.legal_source_id is 'Foreign key to the legal source or precedent cited in the ruling';

-- Person outcome table
comment on table person_outcome is 'Outcome of a ruling as it affects a specific person (e.g., being sanctioned, declared winner). Connects rulings to individuals';
comment on column person_outcome.person_outcome_id is 'Primary key for the person outcome table';
comment on column person_outcome.ruling_id is 'Foreign key linking to the ruling';
comment on column person_outcome.person_id is 'Foreign key linking to the person';
comment on column person_outcome.outcome_type_id is 'Foreign key to the outcome type (e.g., damages, fined, acquitted, winner)';
comment on column person_outcome.description is 'Description of the specific outcome for this person';

-- Person relationship table
comment on table person_relationship is 'Relationships between persons (family, social connections)';
comment on column person_relationship.person_relationship_id is 'Primary key for the person relationship table';
comment on column person_relationship.person_1_id is 'Foreign key to the first person in the relationship';
comment on column person_relationship.person_2_id is 'Foreign key to the second person in the relationship';
comment on column person_relationship.relationship_type_id is 'Foreign key to the type of relationship (e.g., father, mother, sibling, spouse)';
comment on column person_relationship.description is 'Additional description of the relationship';

-- Placename table
comment on table placename is 'Standardized geographical place associated with an entry, linked to the Swedish National Survey (external authority/placename registry) with coordinates';
comment on column placename.placename_id is 'Primary key for the placename table';
comment on column placename.placename is 'The standardized placename (ortnamn)';
comment on column placename.northing is 'Northing coordinate in SWEREF 99 TM (EPSG:3006)';
comment on column placename.easting is 'Easting coordinate in SWEREF 99 TM (EPSG:3006)';
comment on column placename.serial_number is 'Serial number (löpnummer) from the National Survey';
comment on column placename.name_type_code is 'Name type code (namntyp) from the National Survey';
comment on column placename.language_code is 'Language code (språk) from the National Survey';
comment on column placename.parish_code is 'Parish/town code (sockenstad) from the National Survey';
comment on column placename.county_code is 'County code (län) from the National Survey';
comment on column placename.municipality_code is 'Municipality code (kommun) from the National Survey';
comment on column placename.combined_placename is 'Combined placename (combination of multiple name components)';
comment on column placename.parish_name is 'Parish name from the National Survey';
comment on column placename.geom is 'PostGIS geometry point in WGS84 (EPSG:4326) for mapping';

-- Lookup tables
comment on table land_rights_status is 'Description of a person''s legal or customary status to the land as interpreted from the entry (e.g., owned land, no land rights, uncertain)';
comment on column land_rights_status.land_rights_status_id is 'Primary key for the land rights status table';
comment on column land_rights_status.land_rights_status is 'Land rights status name (e.g., ''Ja'', ''Nej'', ''Nja'')';
comment on column land_rights_status.description is 'Description of the land rights status';

comment on table legal_source is 'Normative legal text (law code, regulation, precedent) that a ruling cites or applies';
comment on column legal_source.legal_source_id is 'Primary key for the legal source table';
comment on column legal_source.legal_source_name is 'The name of the legal source or legal precedent cited';

comment on table land_use is 'Categorized description of how land is used or claimed in an entry (e.g., fishing, hunting, herding, reindeer grazing), based on interpretation of the source';
comment on column land_use.land_use_id is 'Primary key for the land use table';
comment on column land_use.description is 'The type of land use or resource (e.g., ''Fishing rights'', ''Reindeer grazing'')';

comment on table parish is 'Lookup table for parishes';
comment on column parish.parish_id is 'Primary key for the parish table';
comment on column parish.parish is 'Parish name - heading comes from the National Survey database of place names';

comment on table outcome_type is 'Categorization describing the kind of decision outcome (e.g., winner, sanction, injunction with fine, partition of land)';
comment on column outcome_type.outcome_type_id is 'Primary key for the outcome type table';
comment on column outcome_type.outcome_type_name is 'Outcome type name (e.g., ''Vinnare'', ''Böter'', ''Friad'')';
comment on column outcome_type.description is 'Description of the outcome type';

comment on table role_type is 'Lookup table for role type categories (social or judicial)';
comment on column role_type.role_type_id is 'Primary key for the role type table';
comment on column role_type.role_type_name is 'Role type category name (''Social'' or ''Juridisk'')';
comment on column role_type.description is 'Description of the role type category';

comment on table role is 'Social or legal role attributed to a person in a specific entry (e.g., Nybyggare, Sámi, plaintiff, defendant). Roles are contextual, not permanent identities';
comment on column role.role_id is 'Primary key for the role table';
comment on column role.role_name is 'Role name (e.g., ''Klagande'', ''Svarande'', ''Vittne'', ''Same'', ''Nybyggare'')';
comment on column role.role_type_id is 'Foreign key to role type (social or judicial)';
comment on column role.description is 'Description of the role';

comment on table season is 'Optional temporal qualifier indicating when the event described in an entry took place (e.g., summer, winter)';
comment on column season.season_id is 'Primary key for the season table';
comment on column season.season_name is 'The name of the season when the disputed resource was primarily used (e.g., ''Winter'', ''Summer'')';

comment on table ruling_type is 'Lookup table for ruling/judgement types';
comment on column ruling_type.ruling_type_id is 'Primary key for the ruling type table';
comment on column ruling_type.ruling_type is 'Ruling type name (e.g., ''Dom'', ''Förlikning'', ''Hänvisning'')';
comment on column ruling_type.description is 'Description of the ruling type';

comment on table source is 'Historical source from which court cases are excerpted (e.g., court records, archival volumes), with identifiers and metadata';
comment on column source.source_id is 'Primary key for the source table';
comment on column source.source_name is 'The full name of the historical source document';
comment on column source.source_abbreviation is 'The abbreviation for the source, used for quick reference (e.g., ''DB'' for court record)';

comment on table relationship_type is 'Lookup table for person relationship types';
comment on column relationship_type.relationship_type_id is 'Primary key for the relationship type table';
comment on column relationship_type.relationship_type_name is 'Relationship type name (e.g., ''Far'', ''Mor'', ''Syskon'', ''Make/Maka'')';
comment on column relationship_type.description is 'Description of the relationship type';

/***********************************************************************************************************
** Procedures Create a single JSON snapshot of the entire schema.
** This is useful for exporting the data for use in a static website or similar.
** select export_schema_as_json('digidiggie_tog')
************************************************************************************************************/

-- drop function if exists export_schema_as_json(text);
create or replace function export_schema_as_json(p_schema text)
returns jsonb
language plpgsql as $$
declare
  r record;
  tables jsonb := '{}'::jsonb;
  rows jsonb;
begin
  for r in
    select tablename
    from pg_catalog.pg_tables
    where schemaname = p_schema
    order by tablename
  loop
    execute format(
      'select coalesce(jsonb_agg(to_jsonb(t) order by 1), ''[]''::jsonb) from %I.%I t',
      p_schema, r.tablename
    ) into rows;
    tables := tables || jsonb_build_object(r.tablename, rows);
  end loop;
  return jsonb_build_object(
    'meta', jsonb_build_object('exported_at', now()),
    'schema', p_schema,
    'tables', tables
  );
end $$;

create or replace procedure sync_sequences()
language plpgsql as $$
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
