
--delete from "properties";

-- NOTE: "properties" and "person_properties" will be removed in new model, but for now we keep them to avoid losing data.

drop table if exists "properties" cascade;
drop table if exists "person_properties" cascade;

create table if not exists "properties" (
    property_id integer primary key,
    property_name text not null,
    description text
);

insert into "properties" ("property_id", "property_name", "description") values
    (1, 'External ID', 'External ID of the individual e.g. from a genealogy database'),
    (2, 'Birth place', 'Birth place of the individual'),
    (3, 'Death place', 'Death place of the individual'),
    (4, 'Residence date', 'Residence date of the individual'),
    (5, 'Event date', 'Date of the event related to the individual'),
    (6, 'Event ID', 'Identifier for the event related to the individual'),
    (7, 'Birth date', 'Birth date (deprecated)'),
    (8, 'Death date', 'Death date (deprecated)');


create table if not exists "person_properties" (
    person_id integer not null,
    property_id integer not null,
    specifier text,
    property_value text,
    primary key (person_id, property_id, specifier),
    foreign key (person_id) references persons(person_id) on delete cascade,
    foreign key (property_id) references properties(property_id) on delete cascade
);

insert into person_properties (person_id, property_id, specifier, property_value) 
  select person_id, 1, 'individual', individual_id::text from persons where coalesce(individual_id::text, '') <> ''
  union all
  select person_id, 1, 'father', father_id::text from persons where coalesce(father_id::text, '') <> ''
  union all
  select person_id, 1, 'mother', mother_id::text from persons where coalesce(mother_id::text, '') <> ''
  union all
  select person_id, 2, 'birth_place', birth_place::text from persons where coalesce(birth_place::text, '') <> ''
  union all
  select person_id, 3, 'death_place', death_place::text from persons where coalesce(death_place::text, '') <> ''
  union all
  select person_id, 4, 'residence_date', residence_date::text from persons where coalesce(residence_date::text, '') <> ''
  union all
  select person_id, 5, 'event_date', event_date::text from persons where coalesce(event_date::text, '') <> ''
  union all
  select person_id, 6, 'event_id', event_id::text from persons where coalesce(event_id::text, '') <> ''
  union all
  select person_id, 7, 'birth_date', birth_date::text from persons where coalesce(birth_date::text, '') <> ''
  union all
  select person_id, 8, 'death_date', death_date::text from persons where coalesce(death_date::text, '') <> ''
  ;

alter table "persons"
    drop column "individual_id",
    drop column "father_id",
    drop column "mother_id",
    drop column "birth_place",
    drop column "death_place",
    drop column "residence_date",
    drop column "event_date",
    drop column "event_id",
    drop column "birth_date",
    drop column "death_date"
;

alter table "persons" drop column "full_name";

alter table "persons"
    add column "full_name" text generated always
      as (coalesce("given_name",'') || ' ' || coalesce("patronymic",'') || ' ' || coalesce("surname",''))
        stored;


