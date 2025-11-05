
SET client_encoding = 'UTF-8';

drop table if exists "legal_sources";
create table if not exists "legal_sources"
 (
    "legal_source_id" serial primary key,
    "legal_source_name" text not null default ('')
);

drop table if exists "parishes";
create table if not exists "parishes"
 (
    "parish_id" serial primary key,
    "parish" text not null default ('')
);

drop table if exists "seasons";
create table if not exists "seasons"
 (
    "season_id" serial primary key,
    "season_name" text not null default ('')
);

drop table if exists "sources";
create table if not exists "sources"
 (
    "source_id" serial primary key,
    "source_name" text not null default (''),
    "source_abbreviation" varchar (255)
);

drop table if exists "land_use";
create table if not exists "land_use"
 (
    "land_use_id" serial primary key,
    "type" text not null default ('')
);

drop table if exists "communities";
create table if not exists "communities"
(
    "community_id" serial primary key,
    "community_name" text not null default (''),
    "parish_id" integer null references "parishes"("parish_id")
);

drop table if exists "winners";
create table if not exists "winners"
 (
    "winner_id" serial primary key,
    "winner_description" text not null default ('')
);

drop table if exists "judgements";
create table if not exists "judgements"
 (
    "judgement_id" serial primary key,
    "sanction" text not null default ('')
);

drop table if exists "persons";
create table if not exists "persons"
 (
    "person_id" serial primary key,
    "individual_id" varchar (255),
    "father_id" varchar (255),
    "mother_id" varchar (255),
    "given_name" text null,
    "patronymic" text null,
    "surname" text null,
    "birth_date" integer null,
    "birth_year" integer null,
    "birth_place" text null,
    "residence_date" integer null,
    "death_date" varchar (255) null,
    "death_year" integer null,
    "death_place" varchar (255) null,
    "event_date" varchar (255) null,
    "event_id" integer null,
    "community_name" text null,
    "full_name" text computed by (coalesce("given_name",'') || ' ' || coalesce("patronymic",'') || ' ' || coalesce("surname",'')) stored
);

drop table if exists "entries";
create table if not exists "entries"
 (
    "entry_id" serial primary key,
    "actor_id" integer null references "persons"("person_id"),
    "community_id" integer null references "communities"("community_id"),
    "year" double precision,
    "description" text,
    "land_rights_status" varchar (255),
    "source_id" integer null references "sources"("source_id"),
    "reference_number" varchar (16),
    "season_id" integer null references "seasons"("season_id"),
    "land_use_id" integer null references "land_use"("land_use_id"),
    "original_placename" varchar (50),
    "winner_id" integer null references "winners"("winner_id"),
    "legal_source_id" integer null references "legal_sources"("legal_source_id"),
    "judgement_id" integer null references "judgements"("judgement_id"),
    "placename_id" integer null references "placenames"("placename_id"),
    "lay_judge_involved" boolean not null default false
);

create index "persons_id_father_idx" on "persons" ("father_id");
create index "persons_id_individu_idx" on "persons" ("individual_id");
create index "persons_id_mother_idx" on "persons" ("mother_id");
create index "entries_reference_number_idx" on "entries" ("reference_number");
create index "parishes_parish_idx" on "parishes" ("parish");
create index "communities_parish_id_idx" ON "communities" ("parish_id");
create index "seasons_season_name_idx" ON "seasons" ("season_name");
create index "entries_actor_id_idx" on "entries" ("actor_id");
create index "entries_community_id_idx" on "entries" ("community_id");
create index "entries_season_id_idx" on "entries" ("season_id");
create index "entries_land_use_id_idx" on "entries" ("land_use_id");
create index "entries_winner_id_idx" on "entries" ("winner_id");
create index "entries_legal_source_id_idx" on "entries" ("legal_source_id");
create index "entries_judgement_id_idx" on "entries" ("judgement_id");
create index "entries_placename_id_idx" on "entries" ("placename_id");
