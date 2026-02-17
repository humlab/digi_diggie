
drop schema if exists digidiggie_tng cascade;
create schema digidiggie_tng;

set search_path  to digidiggie_tng, public;
set role gudrun;

create table "communities" (
  "community_id" serial primary key,
  "community_name" text not null default ''::text,
  "parish_id" int4
);

create table "court_cases" (
  "court_case_id" serial primary key,
  "source_id" int4 not null,
  "reference_number" varchar(16),
  "district_court_name" text,
  "case_date" date,
  "source_text" text
);

create table "entries" (
  "entry_id" serial primary key,
  "court_case_id" int4 not null,
  "year" int4,
  "description" text,
  "season_id" int4,
  "land_use_id" int4,
  "original_placename" varchar(50),
  "placename_id" int4
);

create table "land_right_status" (
  "land_rights_status_id" int4 not null primary key,
  "land_rights_status" varchar(255) not null
);

create table "land_use" (
  "land_use_id" serial primary key,
  "type" text not null default ''::text
);

create table "legal_sources" (
  "legal_source_id" serial primary key,
  "legal_source_name" text not null default ''::text
);

create table "outcome_types" (
  "outcome_type_id" serial primary key,
  "outcome_type_name" text not null,
  "description" text,
  constraint "outcome_types_outcome_type_name_key" unique ("outcome_type_name")
);

create table "parishes" (
  "parish_id" serial primary key,
  "parish" text not null default ''::text
);


create table "person_entries" (
  "person_entry_id" serial primary key,
  "entry_id" int4 not null,
  "actor_id" int4 not null,
  "community_id" int4,
  "land_rights_status_id" int4 not null,
  "role_id" int4,
  "curated_text" text
);

create table "person_outcomes" (
  "person_outcome_id" serial primary key,
  "ruling_id" int4 not null,
  "person_id" int4 not null,
  "outcome_type_id" int4 not null,
  "description" text
);


create table "persons" (
  "person_id" serial primary key,
  "given_name" text,
  "patronymic" text,
  "surname" text,
  "birth_year" int4,
  "death_year" int4,
  "community_name" text,
  "full_name" text generated always as (
        ((((coalesce(given_name, ''::text) || ' '::text) ||
            coalesce(patronymic, ''::text)) || ' '::text) ||
            coalesce(surname, ''::text))
) stored
);


CREATE TABLE  "placenames" (
    "id" serial primary key,
    "ortnamn" text,
    "n" double precision,
    "e" double precision,
    "lopnr" text,
    "namntyp_nr" text,
    "språk_nr" text,
    "sockenstad_nr" text,
    "lan_nr" text,
    "kommun_nr" text,
    "kombo" text,
    "sockenstad" text,
    "nr" text    
);

create table "roles" (
  "role_id" serial primary key,
  "role_name" text not null,
  "description" text,
  constraint "roles_role_name_key" unique ("role_name")
);

create table "ruling_type" (
  "ruling_type_id" int4 not null primary key,
  "ruling_type" varchar(255) not null
);

create table "rulings" (
  "ruling_id" serial primary key,
  "court_case_id" int4 not null,
  "year" int4,
  "description" text,
  "ruling_type_id" int4 not null,
  "legal_source_id" int4,
  constraint "rulings_court_case_id_key" unique ("court_case_id")
);

create table "seasons" (
  "season_id" serial primary key,
  "season_name" text not null default ''::text
);

create table "sources" (
  "source_id" serial primary key,
  "source_name" text not null default ''::text,
  "source_abbreviation" varchar(255)
);


-- Indexes
create index "communities_parish_id_idx" on "communities" using btree (
  "parish_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "court_cases_case_date_idx" on "court_cases" using btree (
  "case_date" "pg_catalog"."date_ops" asc nulls last
);
create index "court_cases_reference_number_idx" on "court_cases" using btree (
  "reference_number" "pg_catalog"."text_ops" asc nulls last
);
create index "court_cases_source_id_idx" on "court_cases" using btree (
  "source_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "entries_reference_number_idx" on "court_cases" using btree (
  "reference_number" "pg_catalog"."text_ops" asc nulls last
);
create index "entries_court_case_id_idx" on "entries" using btree (
  "court_case_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "entries_land_use_id_idx" on "entries" using btree (
  "land_use_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "entries_placename_id_idx" on "entries" using btree (
  "placename_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "rulings_legal_source_id_idx" on "rulings" using btree (
  "legal_source_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "entries_season_id_idx" on "entries" using btree (
  "season_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "parishes_parish_idx" on "parishes" using btree (
  "parish" "pg_catalog"."text_ops" asc nulls last
);
create index "person_entries_actor_id_idx" on "person_entries" using btree (
  "actor_id" "pg_catalog"."int4_ops" asc nulls last
);
create index "person_entries_community_id_idx" on "person_entries" using btree (
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
-- Foreign key constraints

alter table "communities" add constraint "communities_parish_id_fkey" foreign key ("parish_id") references "parishes" ("parish_id") on delete no action on update no action;
alter table "court_cases" add constraint "court_cases_source_id_fkey" foreign key ("source_id") references "sources" ("source_id") on delete no action on update no action;
alter table "entries" add constraint "entries_court_case_id_fkey" foreign key ("court_case_id") references "court_cases" ("court_case_id") on delete cascade on update no action;
alter table "entries" add constraint "entries_land_use_id_fkey" foreign key ("land_use_id") references "land_use" ("land_use_id") on delete no action on update no action;
alter table "entries" add constraint "entries_placename_id_fkey" foreign key ("placename_id") references "placenames" ("fid") on delete no action on update no action;
alter table "entries" add constraint "entries_season_id_fkey" foreign key ("season_id") references "seasons" ("season_id") on delete no action on update no action;
alter table "person_entries" add constraint "person_entries_actor_id_fkey" foreign key ("actor_id") references "persons" ("person_id") on delete no action on update no action;
alter table "person_entries" add constraint "person_entries_community_id_fkey" foreign key ("community_id") references "communities" ("community_id") on delete no action on update no action;
alter table "person_entries" add constraint "person_entries_entry_id_fkey" foreign key ("entry_id") references "entries" ("entry_id") on delete cascade on update no action;
alter table "person_entries" add constraint "person_entries_role_id_fkey" foreign key ("role_id") references "roles" ("role_id") on delete no action on update no action;
alter table "person_entries" add constraint "fk_person_entries_land_right_status_1" foreign key ("land_rights_status_id") references "land_right_status" ("land_rights_status_id");
alter table "person_outcomes" add constraint "person_outcomes_outcome_type_id_fkey" foreign key ("outcome_type_id") references "outcome_types" ("outcome_type_id") on delete no action on update no action;
alter table "person_outcomes" add constraint "person_outcomes_person_id_fkey" foreign key ("person_id") references "persons" ("person_id") on delete no action on update no action;
alter table "person_outcomes" add constraint "person_outcomes_ruling_id_fkey" foreign key ("ruling_id") references "rulings" ("ruling_id") on delete cascade on update no action;
alter table "rulings" add constraint "rulings_court_case_id_fkey" foreign key ("court_case_id") references "court_cases" ("court_case_id") on delete cascade on update no action;
alter table "rulings" add constraint "rulings_legal_source_id_fkey" foreign key ("legal_source_id") references "legal_sources" ("legal_source_id") on delete no action on update no action;
alter table "rulings" add constraint "fk_rulings_ruling_type_1" foreign key ("ruling_type_id") references "ruling_type" ("ruling_type_id");

