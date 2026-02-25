
-- TODO: "properties" and "person_properties" might be removed in new model, but for now we keep them to avoid losing data.

set search_path to "digidiggie_tog";

drop table if exists "person_properties" cascade;
drop table if exists "properties" cascade;

create table if not exists "properties"
(
    "property_id" serial primary key,
    "property_name" text,
    "description" text
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

-- Add function to get table and column metadata, including PK/FK info
CREATE OR REPLACE FUNCTION fn_table_columns(p_schema_name text DEFAULT 'digidiggie_tog'::text) RETURNS TABLE(table_schema information_schema.sql_identifier, table_name information_schema.sql_identifier, column_name information_schema.sql_identifier, ordinal_position information_schema.cardinal_number, data_type information_schema.character_data, numeric_precision information_schema.cardinal_number, numeric_scale information_schema.cardinal_number, character_maximum_length information_schema.cardinal_number, is_nullable information_schema.yes_or_no, is_pk information_schema.yes_or_no, is_fk information_schema.yes_or_no, fk_table_name information_schema.sql_identifier, fk_column_name information_schema.sql_identifier)
    LANGUAGE plpgsql
    AS $$ begin return query with fk_constraint as (
        select distinct fk.conrelid,
            fk.confrelid,
            fk.conkey,
            fk.confrelid::regclass::information_schema.sql_identifier as fk_table_name,
            fkc.attname::information_schema.sql_identifier as fk_column_name
        from pg_constraint as fk
            join pg_attribute fkc on fkc.attrelid = fk.confrelid
            and fkc.attnum = fk.confkey [1]
        where fk.contype = 'f'::char
    )
    select pg_tables.schemaname::information_schema.sql_identifier as table_schema,
        pg_tables.tablename::information_schema.sql_identifier as table_name,
        pg_attribute.attname::information_schema.sql_identifier as column_name,
        pg_attribute.attnum::information_schema.cardinal_number as ordinal_position,
        format_type(pg_attribute.atttypid, null)::information_schema.character_data as data_type,
        case
            pg_attribute.atttypid
            when 21
            /*int2*/
            then 16
            when 23
            /*int4*/
            then 32
            when 20
            /*int8*/
            then 64
            when 1700
            /*numeric*/
            then case
                when pg_attribute.atttypmod = -1 then null
                else ((pg_attribute.atttypmod - 4) >> 16) & 65535 -- calculate the precision
            end
            when 700
            /*float4*/
            then 24
            /*flt_mant_dig*/
            when 701
            /*float8*/
            then 53
            /*dbl_mant_dig*/
            else null
        end::information_schema.cardinal_number as numeric_precision,
        case
            when pg_attribute.atttypid in (21, 23, 20) then 0
            when pg_attribute.atttypid in (1700) then case
                when pg_attribute.atttypmod = -1 then null
                else (pg_attribute.atttypmod - 4) & 65535 -- calculate the scale
            end
            else null
        end::information_schema.cardinal_number as numeric_scale,
        case
            when pg_attribute.atttypid not in (1042, 1043)
            or pg_attribute.atttypmod = -1 then null
            else pg_attribute.atttypmod - 4
        end::information_schema.cardinal_number as character_maximum_length,
        case
            pg_attribute.attnotnull
            when false then 'YES'
            else 'NO'
        end::information_schema.yes_or_no as is_nullable,
        case
            when pk.contype is null then 'NO'
            else 'YES'
        end::information_schema.yes_or_no as is_pk,
        case
            when fk.conrelid is null then 'NO'
            else 'YES'
        end::information_schema.yes_or_no as is_fk,
        fk.fk_table_name,
        fk.fk_column_name
        /*,
        d.column_default::text*/
    from pg_tables
        join pg_class on pg_class.relname = pg_tables.tablename
        join pg_namespace ns on ns.oid = pg_class.relnamespace
        and ns.nspname = pg_tables.schemaname
        join pg_attribute on pg_class.oid = pg_attribute.attrelid
        and pg_attribute.attnum > 0
        left join pg_constraint pk on pk.contype = 'p'::"char"
        and pk.conrelid = pg_class.oid
        and (pg_attribute.attnum = any (pk.conkey))
        left join fk_constraint as fk on fk.conrelid = pg_class.oid
        and (pg_attribute.attnum = any (fk.conkey))
        /*left join information_schema.columns d
        on d.table_schema = pg_tables.schemaname::information_schema.sql_identifier
        and d.table_name = pg_tables.tablename::information_schema.sql_identifier
        and d.column_name = pg_attribute.attname::information_schema.sql_identifier        */
    where true --and pg_tables.tableowner = 'sead_master'
        and pg_attribute.atttypid <> 0::oid
        and pg_tables.schemaname = coalesce(p_schema_name, pg_tables.schemaname)
    order by table_name,
        ordinal_position asc;
    end $$;