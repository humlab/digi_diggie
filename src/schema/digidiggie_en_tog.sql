create schema digidiggie_tog;

set default_tablespace = '';

set default_table_access_method = heap;

create table digidiggie_tog.communities (
    community_id serial not null,
    community_name text default ''::text not null,
    parish_id integer
);

create table digidiggie_tog.entries (
    entry_id serial not null,
    actor_id integer,
    community_id integer,
    year double precision,
    description text,
    land_rights_status character varying(255),
    source_id integer,
    reference_number character varying(16),
    season_id integer,
    land_use_id integer,
    original_placename character varying(50),
    winner_id integer,
    legal_source_id integer,
    judgement_id integer,
    placename_id integer,
    lay_judge_involved boolean default false not null
);

create table digidiggie_tog.judgements (
    judgement_id serial not null,
    sanction text default ''::text not null
);

create table digidiggie_tog.land_use (
    land_use_id serial not null,
    type text default ''::text not null
);

create table digidiggie_tog.legal_sources (
    legal_source_id serial not null,
    legal_source_name text default ''::text not null
);

create table digidiggie_tog.parishes (
    parish_id serial not null,
    parish text default ''::text not null
);

create table digidiggie_tog.person_properties (
    person_property_id serial not null,
    person_id integer,
    property_id integer,
    specifier text,
    property_value text default ''::text not null
);

create table digidiggie_tog.persons (
    person_id serial not null,
    given_name text,
    patronymic text,
    surname text,
    birth_year integer,
    death_year integer,
    community_name text,
    full_name text generated always as (((((coalesce(given_name, ''::text) || ' '::text) || coalesce(patronymic, ''::text)) || ' '::text) || coalesce(surname, ''::text))) stored
);


create table digidiggie_tog.placenames (
    fid integer NOT NULL,
    geom text,
    ortnamn text,
    kvartsruta text,
    nkoordinat double precision,
    ekoordinat double precision,
    lanskod text,
    kommunkod text,
    detaljtyp text,
    sprak text,
    lopnummer text,
    sockenstadkod text,
    sockenstadnamn text,
    geom_point digidiggie_tog.geometry(Point,3006)
);

create table digidiggie_tog.properties (
    property_id serial not null,
    property_name text,
    description text
);

create table digidiggie_tog.seasons (
    season_id serial not null,
    season_name text default ''::text not null
);

create table digidiggie_tog.sources (
    source_id serial not null,
    source_name text default ''::text not null,
    source_abbreviation character varying(255)
);

create table digidiggie_tog.winners (
    winner_id serial not null,
    winner_description text default ''::text not null
);

alter table only digidiggie_tog.communities add constraint communities_pkey primary key (community_id);
alter table only digidiggie_tog.entries add constraint entries_pkey primary key (entry_id);
alter table only digidiggie_tog.judgements add constraint judgements_pkey primary key (judgement_id);
alter table only digidiggie_tog.land_use add constraint land_use_pkey primary key (land_use_id);
alter table only digidiggie_tog.legal_sources add constraint legal_sources_pkey primary key (legal_source_id);
alter table only digidiggie_tog.parishes add constraint parishes_pkey primary key (parish_id);
alter table only digidiggie_tog.person_properties add constraint person_properties_pkey primary key (person_property_id);
alter table only digidiggie_tog.persons add constraint persons_pkey primary key (person_id);
alter table only digidiggie_tog.placenames add constraint placenames_pkey primary key (fid);
alter table only digidiggie_tog.properties add constraint properties_pkey primary key (property_id);
alter table only digidiggie_tog.seasons add constraint seasons_pkey primary key (season_id);
alter table only digidiggie_tog.sources add constraint sources_pkey primary key (source_id);
alter table only digidiggie_tog.winners add constraint winners_pkey primary key (winner_id);

create index communities_parish_id_idx on digidiggie_tog.communities using btree (parish_id);
create index entries_actor_id_idx on digidiggie_tog.entries using btree (actor_id);
create index entries_community_id_idx on digidiggie_tog.entries using btree (community_id);
create index entries_judgement_id_idx on digidiggie_tog.entries using btree (judgement_id);
create index entries_land_use_id_idx on digidiggie_tog.entries using btree (land_use_id);
create index entries_legal_source_id_idx on digidiggie_tog.entries using btree (legal_source_id);
create index entries_placename_id_idx on digidiggie_tog.entries using btree (placename_id);
create index entries_reference_number_idx on digidiggie_tog.entries using btree (reference_number);
create index entries_season_id_idx on digidiggie_tog.entries using btree (season_id);
create index entries_winner_id_idx on digidiggie_tog.entries using btree (winner_id);
create index parishes_parish_idx on digidiggie_tog.parishes using btree (parish);
create index seasons_season_name_idx on digidiggie_tog.seasons using btree (season_name);

alter table only digidiggie_tog.communities
    add constraint communities_parish_id_fkey foreign key (parish_id) references digidiggie_tog.parishes(parish_id);

alter table only digidiggie_tog.entries
    add constraint entries_actor_id_fkey foreign key (actor_id) references digidiggie_tog.persons(person_id);

alter table only digidiggie_tog.entries
    add constraint entries_community_id_fkey foreign key (community_id) references digidiggie_tog.communities(community_id);

alter table only digidiggie_tog.entries
    add constraint entries_judgement_id_fkey foreign key (judgement_id) references digidiggie_tog.judgements(judgement_id);

alter table only digidiggie_tog.entries
    add constraint entries_land_use_id_fkey foreign key (land_use_id) references digidiggie_tog.land_use(land_use_id);

alter table only digidiggie_tog.entries
    add constraint entries_legal_source_id_fkey foreign key (legal_source_id) references digidiggie_tog.legal_sources(legal_source_id);

alter table only digidiggie_tog.entries
    add constraint entries_placename_id_fkey foreign key (placename_id) references digidiggie_tog.placenames(fid);

alter table only digidiggie_tog.entries
    add constraint entries_season_id_fkey foreign key (season_id) references digidiggie_tog.seasons(season_id);

alter table only digidiggie_tog.entries
    add constraint entries_source_id_fkey foreign key (source_id) references digidiggie_tog.sources(source_id);

alter table only digidiggie_tog.entries
    add constraint entries_winner_id_fkey foreign key (winner_id) references digidiggie_tog.winners(winner_id);

alter table only digidiggie_tog.person_properties
    add constraint person_properties_person_id_fkey foreign key (person_id) references digidiggie_tog.persons(person_id);

alter table only digidiggie_tog.person_properties
    add constraint person_properties_property_id_fkey foreign key (property_id) references digidiggie_tog.properties(property_id);


comment on schema digidiggie_tog is 'standard digidiggie The Old Generation schema';
comment on column digidiggie_tog.communities.community_id is 'Primary key for the communities table.';
comment on column digidiggie_tog.communities.community_name is 'The name of the community.';
comment on column digidiggie_tog.communities.parish_id is 'Foreign key linking to the parish this community belongs to.';
comment on column digidiggie_tog.entries.entry_id is 'Primary key for the entries table.';
comment on column digidiggie_tog.entries.actor_id is 'Actor/participant in the event. If multiple actors are involved, each actor gets their own row.';
comment on column digidiggie_tog.entries.community_id is 'Foreign key linking to the community where the event took place.';
comment on column digidiggie_tog.entries.year is 'The year the event occurred, or if unknown, the year the matter was heard at court.';
comment on column digidiggie_tog.entries.description is 'Description of the event in free text. Can be multiple rows per event since each actor involved  has its own row';
comment on column digidiggie_tog.entries.land_rights_status is 'Indicates if the individual had land rights.';
comment on column digidiggie_tog.entries.source_id is 'Foreign key to the source document. Source is specified by abbreviation (e.g., DB=court record).';
comment on column digidiggie_tog.entries.reference_number is 'Reference number within a specific source collection, e.g., K.B. Wiklund''s transcripts.';
comment on column digidiggie_tog.entries.season_id is 'Foreign key to the season when the event occurred.';
comment on column digidiggie_tog.entries.land_use_id is 'Foreign key to the resource or land use type involved in the event.';
comment on column digidiggie_tog.entries.original_placename is 'The place''s name as written in the original source document.';
comment on column digidiggie_tog.entries.winner_id is 'Foreign key to the winning party of the legal case.';
comment on column digidiggie_tog.entries.legal_source_id is 'Foreign key to the legal source or precedent cited.';
comment on column digidiggie_tog.entries.judgement_id is 'Foreign key to the judgement or sanction given.';
comment on column digidiggie_tog.entries.placename_id is 'Foreign key to a standardized placename, from the Swedish National Survey (which enables GIS connection)';
comment on column digidiggie_tog.entries.lay_judge_involved is 'Flag (0/1) indicating if the district court''s board of lay judges (nämnd) played a role.';
comment on column digidiggie_tog.judgements.judgement_id is 'Primary key for the judgements table.';
comment on column digidiggie_tog.judgements.sanction is 'The type of judgement or sanction handed down (e.g., ''Fined'').';
comment on column digidiggie_tog.land_use.land_use_id is 'Primary key for the land use types table.';
comment on column digidiggie_tog.land_use.type is 'The type of land use or resource being discussed (e.g., ''Fishing rights'').';
comment on column digidiggie_tog.legal_sources.legal_source_id is 'Primary key for the legal sources table.';
comment on column digidiggie_tog.legal_sources.legal_source_name is 'The name of the legal source or legal precedent cited.';
comment on column digidiggie_tog.parishes.parish_id is 'Primary key for the parishes table.';
comment on column digidiggie_tog.parishes.parish is 'Heading comes from the National Survey database of place names and means "parish or town", but only parishes are relevant here';
comment on column digidiggie_tog.persons.person_id is 'Primary key for the persons table.';
comment on column digidiggie_tog.persons.given_name is 'The person''s given name(s) (first name).';
comment on column digidiggie_tog.persons.patronymic is 'The person''s patronymic name (e.g., ''Andersson'').';
comment on column digidiggie_tog.persons.surname is 'The person''s surname, byname, or family name.';
comment on column digidiggie_tog.persons.birth_year is 'The year of birth.';
comment on column digidiggie_tog.persons.death_year is 'The year of death.';
comment on column digidiggie_tog.persons.community_name is 'The name of the village where the person primarily resided.';
comment on column digidiggie_tog.persons.full_name is 'The person''s full name, likely a computed or concatenated field.';
comment on column digidiggie_tog.seasons.season_id is 'Primary key for the seasons table.';
comment on column digidiggie_tog.seasons.season_name is 'The name of the season when the disputed resource was primarily used (e.g., ''Winter'', ''Summer'').';
comment on column digidiggie_tog.sources.source_id is 'Primary key for the sources table.';
comment on column digidiggie_tog.sources.source_name is 'The full name of the historical source document.';
comment on column digidiggie_tog.sources.source_abbreviation is 'The abbreviation for the source, used for quick reference.';
comment on column digidiggie_tog.winners.winner_id is 'Primary key for the winners table.';
comment on column digidiggie_tog.winners.winner_description is 'Describes the winning party in a legal case (e.g., ''settler'' or ''reindeer herder'')';

CREATE FUNCTION digidiggie_tog.fn_table_columns(p_schema_name text DEFAULT 'digidiggie_tog'::text) RETURNS TABLE(table_schema information_schema.sql_identifier, table_name information_schema.sql_identifier, column_name information_schema.sql_identifier, ordinal_position information_schema.cardinal_number, data_type information_schema.character_data, numeric_precision information_schema.cardinal_number, numeric_scale information_schema.cardinal_number, character_maximum_length information_schema.cardinal_number, is_nullable information_schema.yes_or_no, is_pk information_schema.yes_or_no, is_fk information_schema.yes_or_no, fk_table_name information_schema.sql_identifier, fk_column_name information_schema.sql_identifier)
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
