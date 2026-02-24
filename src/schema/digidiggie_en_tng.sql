-- Enable PostGIS extension for geometry types in public schema
create extension if not exists postgis schema public;

drop schema if exists digidiggie_tng cascade;
create schema digidiggie_tng;

set search_path  to digidiggie_tng, public;
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

create table entry (
    "entry_id" serial primary key,
    "court_case_id" integer not null,
    "entry_year" integer,
    "curated_text" text,
    "source_placename" text,
    "season_id" integer,
    "land_use_id" integer,
    "placename_id" integer
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

create table person_entry (
    "person_entry_id" serial primary key,
    "entry_id" integer not null,
    "actor_id" integer not null,
    "community_id" integer,
    "land_rights_status_id" integer not null,
    "role_id" integer,
    "curated_text" text
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
    "geom" geometry(Point, 4326)        --> WGS84 (EPSG:4326)
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
    "season_name" text not null default ''::text
);

create table ruling_type (
    "ruling_type_id" serial primary key,
    "ruling_type" varchar(255) not null
);

create table ruling (
    "ruling_id" serial primary key,
    "court_case_id" integer not null,
    "year" integer,
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

/***********************************************************************************************************
** Foreign key constraints
************************************************************************************************************/

alter table "community" add constraint "community_parish_id_fkey" foreign key ("parish_id") references "parish" ("parish_id") on delete no action on update no action;
alter table "court_case" add constraint "court_case_source_id_fkey" foreign key ("source_id") references "source" ("source_id") on delete no action on update no action;
alter table "entry" add constraint "entry_court_case_id_fkey" foreign key ("court_case_id") references "court_case" ("court_case_id") on delete cascade on update no action;
alter table "entry" add constraint "entry_land_use_id_fkey" foreign key ("land_use_id") references "land_use" ("land_use_id") on delete no action on update no action;
-- FIXME: #13 Update to constraint to reference correct column in `placename`.
-- alter table "entry" add constraint "entry_placename_id_fkey" foreign key ("placename_id") references "placename" ("fid") on delete no action on update no action;
alter table "entry" add constraint "entry_season_id_fkey" foreign key ("season_id") references "season" ("season_id") on delete no action on update no action;
alter table "person_entry" add constraint "person_entry_actor_id_fkey" foreign key ("actor_id") references "person" ("person_id") on delete no action on update no action;
alter table "person_entry" add constraint "person_entry_community_id_fkey" foreign key ("community_id") references "community" ("community_id") on delete no action on update no action;
alter table "person_entry" add constraint "person_entry_entry_id_fkey" foreign key ("entry_id") references "entry" ("entry_id") on delete cascade on update no action;
alter table "person_entry" add constraint "person_entry_role_id_fkey" foreign key ("role_id") references "role" ("role_id") on delete no action on update no action;
alter table "person_entry" add constraint "fk_person_entry_land_rights_status_1" foreign key ("land_rights_status_id") references "land_rights_status" ("land_rights_status_id");
alter table "person_outcomes" add constraint "person_outcomes_outcome_type_id_fkey" foreign key ("outcome_type_id") references "outcome_types" ("outcome_type_id") on delete no action on update no action;
alter table "person_outcomes" add constraint "person_outcomes_person_id_fkey" foreign key ("person_id") references "person" ("person_id") on delete no action on update no action;
alter table "person_outcomes" add constraint "person_outcomes_ruling_id_fkey" foreign key ("ruling_id") references "rulings" ("ruling_id") on delete cascade on update no action;
alter table "rulings" add constraint "rulings_court_case_id_fkey" foreign key ("court_case_id") references "court_cases" ("court_case_id") on delete cascade on update no action;
alter table "rulings" add constraint "rulings_legal_source_id_fkey" foreign key ("legal_source_id") references "legal_sources" ("legal_source_id") on delete no action on update no action;
alter table "rulings" add constraint "fk_rulings_ruling_type_1" foreign key ("ruling_type_id") references "ruling_type" ("ruling_type_id");


/***********************************************************************************************************
** Indexes
************************************************************************************************************/

create index "communities_parish_id_idx" on "community" using btree (
  "parish_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "court_cases_case_year_idx" on "court_cases" using btree (
  "case_year" "pg_catalog"."int4_ops" asc nulls last
);
create index "court_cases_reference_number_idx" on "court_case" using btree (
  "reference_number" "pg_catalog"."text_ops" asc nulls last
);
create index "court_cases_source_id_idx" on "court_case" using btree (
  "source_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "entries_reference_number_idx" on "entry" using btree (
  "reference_number" "pg_catalog"."text_ops" asc nulls last
);
create index "entries_court_case_id_idx" on "entry" using btree (
  "court_case_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "entries_land_use_id_idx" on "entry" using btree (
  "land_use_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "entries_placename_id_idx" on "entry" using btree (
  "placename_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "rulings_legal_source_id_idx" on "rulings" using btree (
  "legal_source_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "entries_season_id_idx" on "entry" using btree (
  "season_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "parishes_parish_idx" on "parish" using btree (
  "parish" "pg_catalog"."text_ops" asc nulls last
);
create index "person_entries_actor_id_idx" on "person_entry" using btree (
  "actor_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "person_entries_community_id_idx" on "person_entry" using btree (
  "community_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "person_entries_entry_id_idx" on "person_entries" using btree (
  "entry_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "person_outcomes_outcome_type_id_idx" on "person_outcomes" using btree (
  "outcome_type_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "person_outcomes_person_id_idx" on "person_outcomes" using btree (
  "person_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "person_outcomes_ruling_id_idx" on "person_outcomes" using btree (
  "ruling_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "seasons_season_name_idx" on "seasons" using btree (
  "season_name" "pg_catalog"."text_ops" asc nulls last
);
