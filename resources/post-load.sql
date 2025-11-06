
delete from "properties";

insert into "properties" ("property_id", "property_name", "Description") values
    ('External ID', 'External ID of an individual e.g. from a genealogy database'),
    ('Birth place', 'Birth place of an individual'),
    ('Death place', 'Death place of an individual'),
    ('Residence data', 'Residence data of an individual');

insert into "person_properties" ("person_id", "property_id", "specifier", "property_value") 
    select p."person_id", pr."property_id", 'individual', p."individual_id"::text
    from "persons" p
    join "properties" pr on pr."property_name" = 'External ID'
    where pr."property_name" = 'External ID'
      and coalesce(p."individual_id", '') <> ''
    union all
    select p."person_id", pr."property_id", 'father', p."father_id"::text
    from "persons" p
    join "properties" pr on pr."property_name" = 'Father (external ID)'
    where pr."property_name" = 'Father (external ID)'
      and coalesce(p."father_id", '') <> ''
    union all
    select p."person_id", pr."property_id", p."mother_id"::text
    from "persons" p
    join "properties" pr on pr."property_name" = 'Mother (external ID)'
    where pr."property_name" = 'Mother (external ID)'
        and coalesce(p."mother_id", '') <> ''
    union all
    select p."person_id", pr."property_id", p."birth_place"::text
    from "persons" p
    join "properties" pr on pr."property_name" = 'Birth place'
    where pr."property_name" = 'Birth place'
      and coalesce(p."birth_place", '') <> ''
    union all
    select p."person_id", pr."property_id", p."death_place"::text
    from "persons" p
    join "properties" pr on pr."property_name" = 'Death place'
    where pr."property_name" = 'Death place'
      and coalesce(p."death_place", '') <> '';

alter  table "persons"
    drop column "individual_id",
    drop column "father_id",
    drop column "mother_id"
    drop column "birth_place",
    drop column "death_place";
