--
-- PostgreSQL database dump
--

-- Dumped from database version 17.4 (Debian 17.4-1.pgdg110+2)
-- Dumped by pg_dump version 18.2

SET statement_timeout = 0;
SET lock_timeout = 0;
SET idle_in_transaction_session_timeout = 0;
SET transaction_timeout = 0;
SET client_encoding = 'UTF8';
SET standard_conforming_strings = on;
SELECT pg_catalog.set_config('search_path', '', false);
SET check_function_bodies = false;
SET xmloption = content;
SET client_min_messages = warning;
SET row_security = off;

--
-- Name: digidiggie_tog; Type: SCHEMA; Schema: -; Owner: pg_database_owner
--

CREATE SCHEMA digidiggie_tog;


ALTER SCHEMA digidiggie_tog OWNER TO pg_database_owner;

--
-- Name: SCHEMA digidiggie_tog; Type: COMMENT; Schema: -; Owner: pg_database_owner
--

COMMENT ON SCHEMA digidiggie_tog IS 'standard digidiggie The Old Generation schema';


--
-- Name: fn_table_columns(text); Type: FUNCTION; Schema: digidiggie_tog; Owner: gudrun
--

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


ALTER FUNCTION digidiggie_tog.fn_table_columns(p_schema_name text) OWNER TO gudrun;

SET default_tablespace = '';

SET default_table_access_method = heap;

--
-- Name: communities; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.communities (
    community_id integer NOT NULL,
    community_name text DEFAULT ''::text NOT NULL,
    parish_id integer
);


ALTER TABLE digidiggie_tog.communities OWNER TO gudrun;

--
-- Name: COLUMN communities.community_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.communities.community_id IS 'Primary key for the communities table.';


--
-- Name: COLUMN communities.community_name; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.communities.community_name IS 'The name of the community.';


--
-- Name: COLUMN communities.parish_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.communities.parish_id IS 'Foreign key linking to the parish this community belongs to.';


--
-- Name: communities_community_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.communities_community_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.communities_community_id_seq OWNER TO gudrun;

--
-- Name: communities_community_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.communities_community_id_seq OWNED BY digidiggie_tog.communities.community_id;


--
-- Name: entries; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.entries (
    entry_id integer NOT NULL,
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
    lay_judge_involved boolean DEFAULT false NOT NULL
);


ALTER TABLE digidiggie_tog.entries OWNER TO gudrun;

--
-- Name: COLUMN entries.entry_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.entry_id IS 'Primary key for the entries table.';


--
-- Name: COLUMN entries.actor_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.actor_id IS 'Actor/participant in the event. If multiple actors are involved, each actor gets their own row.';


--
-- Name: COLUMN entries.community_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.community_id IS 'Foreign key linking to the community where the event took place.';


--
-- Name: COLUMN entries.year; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.year IS 'The year the event occurred, or if unknown, the year the matter was heard at court.';


--
-- Name: COLUMN entries.description; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.description IS 'Description of the event in free text. Can be multiple rows per event since each actor involved  has its own row';


--
-- Name: COLUMN entries.land_rights_status; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.land_rights_status IS 'Indicates if the individual had land rights.';


--
-- Name: COLUMN entries.source_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.source_id IS 'Foreign key to the source document. Source is specified by abbreviation (e.g., DB=court record).';


--
-- Name: COLUMN entries.reference_number; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.reference_number IS 'Reference number within a specific source collection, e.g., K.B. Wiklund''s transcripts.';


--
-- Name: COLUMN entries.season_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.season_id IS 'Foreign key to the season when the event occurred.';


--
-- Name: COLUMN entries.land_use_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.land_use_id IS 'Foreign key to the resource or land use type involved in the event.';


--
-- Name: COLUMN entries.original_placename; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.original_placename IS 'The place''s name as written in the original source document.';


--
-- Name: COLUMN entries.winner_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.winner_id IS 'Foreign key to the winning party of the legal case.';


--
-- Name: COLUMN entries.legal_source_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.legal_source_id IS 'Foreign key to the legal source or precedent cited.';


--
-- Name: COLUMN entries.judgement_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.judgement_id IS 'Foreign key to the judgement or sanction given.';


--
-- Name: COLUMN entries.placename_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.placename_id IS 'Foreign key to a standardized placename, from the Swedish National Survey (which enables GIS connection)';


--
-- Name: COLUMN entries.lay_judge_involved; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.entries.lay_judge_involved IS 'Flag (0/1) indicating if the district court''s board of lay judges (nämnd) played a role.';


--
-- Name: entries_entry_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.entries_entry_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.entries_entry_id_seq OWNER TO gudrun;

--
-- Name: entries_entry_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.entries_entry_id_seq OWNED BY digidiggie_tog.entries.entry_id;


--
-- Name: judgements; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.judgements (
    judgement_id integer NOT NULL,
    sanction text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie_tog.judgements OWNER TO gudrun;

--
-- Name: COLUMN judgements.judgement_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.judgements.judgement_id IS 'Primary key for the judgements table.';


--
-- Name: COLUMN judgements.sanction; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.judgements.sanction IS 'The type of judgement or sanction handed down (e.g., ''Fined'').';


--
-- Name: judgements_judgement_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.judgements_judgement_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.judgements_judgement_id_seq OWNER TO gudrun;

--
-- Name: judgements_judgement_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.judgements_judgement_id_seq OWNED BY digidiggie_tog.judgements.judgement_id;


--
-- Name: land_use; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.land_use (
    land_use_id integer NOT NULL,
    type text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie_tog.land_use OWNER TO gudrun;

--
-- Name: COLUMN land_use.land_use_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.land_use.land_use_id IS 'Primary key for the land use types table.';


--
-- Name: COLUMN land_use.type; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.land_use.type IS 'The type of land use or resource being discussed (e.g., ''Fishing rights'').';


--
-- Name: land_use_land_use_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.land_use_land_use_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.land_use_land_use_id_seq OWNER TO gudrun;

--
-- Name: land_use_land_use_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.land_use_land_use_id_seq OWNED BY digidiggie_tog.land_use.land_use_id;


--
-- Name: legal_sources; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.legal_sources (
    legal_source_id integer NOT NULL,
    legal_source_name text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie_tog.legal_sources OWNER TO gudrun;

--
-- Name: COLUMN legal_sources.legal_source_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.legal_sources.legal_source_id IS 'Primary key for the legal sources table.';


--
-- Name: COLUMN legal_sources.legal_source_name; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.legal_sources.legal_source_name IS 'The name of the legal source or legal precedent cited.';


--
-- Name: legal_sources_legal_source_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.legal_sources_legal_source_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.legal_sources_legal_source_id_seq OWNER TO gudrun;

--
-- Name: legal_sources_legal_source_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.legal_sources_legal_source_id_seq OWNED BY digidiggie_tog.legal_sources.legal_source_id;


--
-- Name: parishes; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.parishes (
    parish_id integer NOT NULL,
    parish text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie_tog.parishes OWNER TO gudrun;

--
-- Name: COLUMN parishes.parish_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.parishes.parish_id IS 'Primary key for the parishes table.';


--
-- Name: COLUMN parishes.parish; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.parishes.parish IS 'Heading comes from the National Survey database of place names and means "parish or town", but only parishes are relevant here';


--
-- Name: parishes_parish_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.parishes_parish_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.parishes_parish_id_seq OWNER TO gudrun;

--
-- Name: parishes_parish_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.parishes_parish_id_seq OWNED BY digidiggie_tog.parishes.parish_id;


--
-- Name: person_properties; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.person_properties (
    person_property_id integer NOT NULL,
    person_id integer,
    property_id integer,
    specifier text,
    property_value text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie_tog.person_properties OWNER TO gudrun;

--
-- Name: person_properties_person_property_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.person_properties_person_property_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.person_properties_person_property_id_seq OWNER TO gudrun;

--
-- Name: person_properties_person_property_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.person_properties_person_property_id_seq OWNED BY digidiggie_tog.person_properties.person_property_id;


--
-- Name: persons; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.persons (
    person_id integer NOT NULL,
    given_name text,
    patronymic text,
    surname text,
    birth_year integer,
    death_year integer,
    community_name text,
    full_name text GENERATED ALWAYS AS (((((COALESCE(given_name, ''::text) || ' '::text) || COALESCE(patronymic, ''::text)) || ' '::text) || COALESCE(surname, ''::text))) STORED
);


ALTER TABLE digidiggie_tog.persons OWNER TO gudrun;

--
-- Name: COLUMN persons.person_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.persons.person_id IS 'Primary key for the persons table.';


--
-- Name: COLUMN persons.given_name; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.persons.given_name IS 'The person''s given name(s) (first name).';


--
-- Name: COLUMN persons.patronymic; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.persons.patronymic IS 'The person''s patronymic name (e.g., ''Andersson'').';


--
-- Name: COLUMN persons.surname; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.persons.surname IS 'The person''s surname, byname, or family name.';


--
-- Name: COLUMN persons.birth_year; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.persons.birth_year IS 'The year of birth.';


--
-- Name: COLUMN persons.death_year; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.persons.death_year IS 'The year of death.';


--
-- Name: COLUMN persons.community_name; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.persons.community_name IS 'The name of the village where the person primarily resided.';


--
-- Name: COLUMN persons.full_name; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.persons.full_name IS 'The person''s full name, likely a computed or concatenated field.';


--
-- Name: persons_person_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.persons_person_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.persons_person_id_seq OWNER TO gudrun;

--
-- Name: persons_person_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.persons_person_id_seq OWNED BY digidiggie_tog.persons.person_id;


--
-- Name: placenames; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.placenames (
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


ALTER TABLE digidiggie_tog.placenames OWNER TO gudrun;

--
-- Name: properties; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.properties (
    property_id integer NOT NULL,
    property_name text,
    description text
);


ALTER TABLE digidiggie_tog.properties OWNER TO gudrun;

--
-- Name: properties_property_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.properties_property_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.properties_property_id_seq OWNER TO gudrun;

--
-- Name: properties_property_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.properties_property_id_seq OWNED BY digidiggie_tog.properties.property_id;


--
-- Name: seasons; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.seasons (
    season_id integer NOT NULL,
    season_name text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie_tog.seasons OWNER TO gudrun;

--
-- Name: COLUMN seasons.season_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.seasons.season_id IS 'Primary key for the seasons table.';


--
-- Name: COLUMN seasons.season_name; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.seasons.season_name IS 'The name of the season when the disputed resource was primarily used (e.g., ''Winter'', ''Summer'').';


--
-- Name: seasons_season_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.seasons_season_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.seasons_season_id_seq OWNER TO gudrun;

--
-- Name: seasons_season_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.seasons_season_id_seq OWNED BY digidiggie_tog.seasons.season_id;


--
-- Name: sources; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.sources (
    source_id integer NOT NULL,
    source_name text DEFAULT ''::text NOT NULL,
    source_abbreviation character varying(255)
);


ALTER TABLE digidiggie_tog.sources OWNER TO gudrun;

--
-- Name: COLUMN sources.source_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.sources.source_id IS 'Primary key for the sources table.';


--
-- Name: COLUMN sources.source_name; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.sources.source_name IS 'The full name of the historical source document.';


--
-- Name: COLUMN sources.source_abbreviation; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.sources.source_abbreviation IS 'The abbreviation for the source, used for quick reference.';


--
-- Name: sources_source_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.sources_source_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.sources_source_id_seq OWNER TO gudrun;

--
-- Name: sources_source_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.sources_source_id_seq OWNED BY digidiggie_tog.sources.source_id;


--
-- Name: winners; Type: TABLE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE TABLE digidiggie_tog.winners (
    winner_id integer NOT NULL,
    winner_description text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie_tog.winners OWNER TO gudrun;

--
-- Name: COLUMN winners.winner_id; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.winners.winner_id IS 'Primary key for the winners table.';


--
-- Name: COLUMN winners.winner_description; Type: COMMENT; Schema: digidiggie_tog; Owner: gudrun
--

COMMENT ON COLUMN digidiggie_tog.winners.winner_description IS 'Describes the winning party in a legal case (e.g., ''settler'' or ''reindeer herder'')';


--
-- Name: winners_winner_id_seq; Type: SEQUENCE; Schema: digidiggie_tog; Owner: gudrun
--

CREATE SEQUENCE digidiggie_tog.winners_winner_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie_tog.winners_winner_id_seq OWNER TO gudrun;

--
-- Name: winners_winner_id_seq; Type: SEQUENCE OWNED BY; Schema: digidiggie_tog; Owner: gudrun
--

ALTER SEQUENCE digidiggie_tog.winners_winner_id_seq OWNED BY digidiggie_tog.winners.winner_id;


--
-- Name: communities community_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.communities ALTER COLUMN community_id SET DEFAULT nextval('digidiggie_tog.communities_community_id_seq'::regclass);


--
-- Name: entries entry_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries ALTER COLUMN entry_id SET DEFAULT nextval('digidiggie_tog.entries_entry_id_seq'::regclass);


--
-- Name: judgements judgement_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.judgements ALTER COLUMN judgement_id SET DEFAULT nextval('digidiggie_tog.judgements_judgement_id_seq'::regclass);


--
-- Name: land_use land_use_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.land_use ALTER COLUMN land_use_id SET DEFAULT nextval('digidiggie_tog.land_use_land_use_id_seq'::regclass);


--
-- Name: legal_sources legal_source_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.legal_sources ALTER COLUMN legal_source_id SET DEFAULT nextval('digidiggie_tog.legal_sources_legal_source_id_seq'::regclass);


--
-- Name: parishes parish_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.parishes ALTER COLUMN parish_id SET DEFAULT nextval('digidiggie_tog.parishes_parish_id_seq'::regclass);


--
-- Name: person_properties person_property_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.person_properties ALTER COLUMN person_property_id SET DEFAULT nextval('digidiggie_tog.person_properties_person_property_id_seq'::regclass);


--
-- Name: persons person_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.persons ALTER COLUMN person_id SET DEFAULT nextval('digidiggie_tog.persons_person_id_seq'::regclass);


--
-- Name: properties property_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.properties ALTER COLUMN property_id SET DEFAULT nextval('digidiggie_tog.properties_property_id_seq'::regclass);


--
-- Name: seasons season_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.seasons ALTER COLUMN season_id SET DEFAULT nextval('digidiggie_tog.seasons_season_id_seq'::regclass);


--
-- Name: sources source_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.sources ALTER COLUMN source_id SET DEFAULT nextval('digidiggie_tog.sources_source_id_seq'::regclass);


--
-- Name: winners winner_id; Type: DEFAULT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.winners ALTER COLUMN winner_id SET DEFAULT nextval('digidiggie_tog.winners_winner_id_seq'::regclass);


--
-- Name: communities communities_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.communities
    ADD CONSTRAINT communities_pkey PRIMARY KEY (community_id);


--
-- Name: entries entries_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries
    ADD CONSTRAINT entries_pkey PRIMARY KEY (entry_id);


--
-- Name: judgements judgements_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.judgements
    ADD CONSTRAINT judgements_pkey PRIMARY KEY (judgement_id);


--
-- Name: land_use land_use_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.land_use
    ADD CONSTRAINT land_use_pkey PRIMARY KEY (land_use_id);


--
-- Name: legal_sources legal_sources_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.legal_sources
    ADD CONSTRAINT legal_sources_pkey PRIMARY KEY (legal_source_id);


--
-- Name: parishes parishes_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.parishes
    ADD CONSTRAINT parishes_pkey PRIMARY KEY (parish_id);


--
-- Name: person_properties person_properties_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.person_properties
    ADD CONSTRAINT person_properties_pkey PRIMARY KEY (person_property_id);


--
-- Name: persons persons_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.persons
    ADD CONSTRAINT persons_pkey PRIMARY KEY (person_id);


--
-- Name: placenames placenames_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.placenames
    ADD CONSTRAINT placenames_pkey PRIMARY KEY (fid);


--
-- Name: properties properties_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.properties
    ADD CONSTRAINT properties_pkey PRIMARY KEY (property_id);


--
-- Name: seasons seasons_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.seasons
    ADD CONSTRAINT seasons_pkey PRIMARY KEY (season_id);


--
-- Name: sources sources_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.sources
    ADD CONSTRAINT sources_pkey PRIMARY KEY (source_id);


--
-- Name: winners winners_pkey; Type: CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.winners
    ADD CONSTRAINT winners_pkey PRIMARY KEY (winner_id);


--
-- Name: communities_parish_id_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX communities_parish_id_idx ON digidiggie_tog.communities USING btree (parish_id);


--
-- Name: entries_actor_id_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX entries_actor_id_idx ON digidiggie_tog.entries USING btree (actor_id);


--
-- Name: entries_community_id_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX entries_community_id_idx ON digidiggie_tog.entries USING btree (community_id);


--
-- Name: entries_judgement_id_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX entries_judgement_id_idx ON digidiggie_tog.entries USING btree (judgement_id);


--
-- Name: entries_land_use_id_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX entries_land_use_id_idx ON digidiggie_tog.entries USING btree (land_use_id);


--
-- Name: entries_legal_source_id_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX entries_legal_source_id_idx ON digidiggie_tog.entries USING btree (legal_source_id);


--
-- Name: entries_placename_id_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX entries_placename_id_idx ON digidiggie_tog.entries USING btree (placename_id);


--
-- Name: entries_reference_number_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX entries_reference_number_idx ON digidiggie_tog.entries USING btree (reference_number);


--
-- Name: entries_season_id_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX entries_season_id_idx ON digidiggie_tog.entries USING btree (season_id);


--
-- Name: entries_winner_id_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX entries_winner_id_idx ON digidiggie_tog.entries USING btree (winner_id);


--
-- Name: parishes_parish_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX parishes_parish_idx ON digidiggie_tog.parishes USING btree (parish);


--
-- Name: seasons_season_name_idx; Type: INDEX; Schema: digidiggie_tog; Owner: gudrun
--

CREATE INDEX seasons_season_name_idx ON digidiggie_tog.seasons USING btree (season_name);


--
-- Name: communities communities_parish_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.communities
    ADD CONSTRAINT communities_parish_id_fkey FOREIGN KEY (parish_id) REFERENCES digidiggie_tog.parishes(parish_id);


--
-- Name: entries entries_actor_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries
    ADD CONSTRAINT entries_actor_id_fkey FOREIGN KEY (actor_id) REFERENCES digidiggie_tog.persons(person_id);


--
-- Name: entries entries_community_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries
    ADD CONSTRAINT entries_community_id_fkey FOREIGN KEY (community_id) REFERENCES digidiggie_tog.communities(community_id);


--
-- Name: entries entries_judgement_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries
    ADD CONSTRAINT entries_judgement_id_fkey FOREIGN KEY (judgement_id) REFERENCES digidiggie_tog.judgements(judgement_id);


--
-- Name: entries entries_land_use_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries
    ADD CONSTRAINT entries_land_use_id_fkey FOREIGN KEY (land_use_id) REFERENCES digidiggie_tog.land_use(land_use_id);


--
-- Name: entries entries_legal_source_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries
    ADD CONSTRAINT entries_legal_source_id_fkey FOREIGN KEY (legal_source_id) REFERENCES digidiggie_tog.legal_sources(legal_source_id);


--
-- Name: entries entries_placename_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries
    ADD CONSTRAINT entries_placename_id_fkey FOREIGN KEY (placename_id) REFERENCES digidiggie_tog.placenames(fid);


--
-- Name: entries entries_season_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries
    ADD CONSTRAINT entries_season_id_fkey FOREIGN KEY (season_id) REFERENCES digidiggie_tog.seasons(season_id);


--
-- Name: entries entries_source_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries
    ADD CONSTRAINT entries_source_id_fkey FOREIGN KEY (source_id) REFERENCES digidiggie_tog.sources(source_id);


--
-- Name: entries entries_winner_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.entries
    ADD CONSTRAINT entries_winner_id_fkey FOREIGN KEY (winner_id) REFERENCES digidiggie_tog.winners(winner_id);


--
-- Name: person_properties person_properties_person_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.person_properties
    ADD CONSTRAINT person_properties_person_id_fkey FOREIGN KEY (person_id) REFERENCES digidiggie_tog.persons(person_id);


--
-- Name: person_properties person_properties_property_id_fkey; Type: FK CONSTRAINT; Schema: digidiggie_tog; Owner: gudrun
--

ALTER TABLE ONLY digidiggie_tog.person_properties
    ADD CONSTRAINT person_properties_property_id_fkey FOREIGN KEY (property_id) REFERENCES digidiggie_tog.properties(property_id);


--
-- PostgreSQL database dump complete
--

\unrestrict rf8aaibpfEEUsBZCx0MD7GeLGVbQDplNu9rBHCMoWxyqJfI4nYyOhxDwAbq4kWU

