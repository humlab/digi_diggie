
-- delete from "properties";

-- insert into "properties" ("property_id", "property_name", "description") values
--     (1, 'External ID', 'External ID of the individual e.g. from a genealogy database'),
--     (2, 'Birth place', 'Birth place of the individual'),
--     (3, 'Death place', 'Death place of the individual'),
--     (4, 'Residence date', 'Residence date of the individual'),
--     (5, 'Event date', 'Date of the event related to the individual'),
--     (6, 'Event ID', 'Identifier for the event related to the individual'),
--     (7, 'Birth date', 'Birth date (deprecated)'),
--     (8, 'Death date', 'Death date (deprecated)');

-- # FIXME: Should `individual_id`, `father_id`, and `mother_id` all have `property_id` = 1?
-- # TODO: Add specifiers
-- # NOTE: This will be removed in new model
-- insert into person_properties (person_id, property_id, specifier, property_value) 
--   select person_id, 1, 'individual', individual_id::text from persons where coalesce(individual_id::text, '') <> ''
--   union all
--   select person_id, 1, 'father', father_id::text from persons where coalesce(father_id::text, '') <> ''
--   union all
--   select person_id, 1, 'mother', mother_id::text from persons where coalesce(mother_id::text, '') <> ''
--   union all
--   select person_id, 2, null, birth_place::text from persons where coalesce(birth_place::text, '') <> ''
--   union all
--   select person_id, 3, null, death_place::text from persons where coalesce(death_place::text, '') <> ''
--   union all
--   select person_id, 4, null, residence_date::text from persons where coalesce(residence_date::text, '') <> ''
--   union all
--   select person_id, 5, null, event_date::text from persons where coalesce(event_date::text, '') <> ''
--   union all
--   select person_id, 6, null, event_id::text from persons where coalesce(event_id::text, '') <> ''
--   union all
--   select person_id, 7, null, birth_date::text from persons where coalesce(birth_date::text, '') <> ''
--   union all
--   select person_id, 8, null, death_date::text from persons where coalesce(death_date::text, '') <> ''
--   ;

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


