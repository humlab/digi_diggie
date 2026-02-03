--
-- PostgreSQL database dump

create schema if not exists digidiggie;

--
-- Name: communities; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.communities (
    community_id integer NOT NULL,
    community_name text DEFAULT ''::text NOT NULL,
    parish_id integer
);


ALTER TABLE digidiggie.communities OWNER TO gudrun;

--
-- Name: COLUMN communities.community_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.communities.community_id IS 'Primary key for the communities table.';


--
-- Name: COLUMN communities.community_name; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.communities.community_name IS 'The name of the community.';


--
-- Name: COLUMN communities.parish_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.communities.parish_id IS 'Foreign key linking to the parish this community belongs to.';


--
-- Name: communities_community_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.communities_community_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.communities_community_id_seq OWNER TO gudrun;

--
-- Name: communities_community_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.communities_community_id_seq OWNED BY digidiggie.communities.community_id;


--
-- Name: entries; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.entries (
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


ALTER TABLE digidiggie.entries OWNER TO gudrun;

--
-- Name: COLUMN entries.entry_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.entry_id IS 'Primary key for the entries table.';


--
-- Name: COLUMN entries.actor_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.actor_id IS 'Actor/participant in the event. If multiple actors are involved, each actor gets their own row.';


--
-- Name: COLUMN entries.community_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.community_id IS 'Foreign key linking to the community where the event took place.';


--
-- Name: COLUMN entries.year; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.year IS 'The year the event occurred, or if unknown, the year the matter was heard at court.';


--
-- Name: COLUMN entries.description; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.description IS 'Description of the event in free text. Can be multiple rows per event since each actor involved  has its own row';


--
-- Name: COLUMN entries.land_rights_status; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.land_rights_status IS 'Indicates if the individual had land rights.';


--
-- Name: COLUMN entries.source_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.source_id IS 'Foreign key to the source document. Source is specified by abbreviation (e.g., DB=court record).';


--
-- Name: COLUMN entries.reference_number; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.reference_number IS 'Reference number within a specific source collection, e.g., K.B. Wiklund''s transcripts.';


--
-- Name: COLUMN entries.season_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.season_id IS 'Foreign key to the season when the event occurred.';


--
-- Name: COLUMN entries.land_use_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.land_use_id IS 'Foreign key to the resource or land use type involved in the event.';


--
-- Name: COLUMN entries.original_placename; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.original_placename IS 'The place''s name as written in the original source document.';


--
-- Name: COLUMN entries.winner_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.winner_id IS 'Foreign key to the winning party of the legal case.';


--
-- Name: COLUMN entries.legal_source_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.legal_source_id IS 'Foreign key to the legal source or precedent cited.';


--
-- Name: COLUMN entries.judgement_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.judgement_id IS 'Foreign key to the judgement or sanction given.';


--
-- Name: COLUMN entries.placename_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.placename_id IS 'Foreign key to a standardized placename, from the Swedish National Survey (which enables GIS connection)';


--
-- Name: COLUMN entries.lay_judge_involved; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.entries.lay_judge_involved IS 'Flag (0/1) indicating if the district court''s board of lay judges (nämnd) played a role.';


--
-- Name: entries_entry_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.entries_entry_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.entries_entry_id_seq OWNER TO gudrun;

--
-- Name: entries_entry_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.entries_entry_id_seq OWNED BY digidiggie.entries.entry_id;


--
-- Name: judgements; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.judgements (
    judgement_id integer NOT NULL,
    sanction text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie.judgements OWNER TO gudrun;

--
-- Name: COLUMN judgements.judgement_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.judgements.judgement_id IS 'Primary key for the judgements table.';


--
-- Name: COLUMN judgements.sanction; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.judgements.sanction IS 'The type of judgement or sanction handed down (e.g., ''Fined'').';


--
-- Name: judgements_judgement_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.judgements_judgement_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.judgements_judgement_id_seq OWNER TO gudrun;

--
-- Name: judgements_judgement_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.judgements_judgement_id_seq OWNED BY digidiggie.judgements.judgement_id;


--
-- Name: land_use; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.land_use (
    land_use_id integer NOT NULL,
    type text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie.land_use OWNER TO gudrun;

--
-- Name: COLUMN land_use.land_use_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.land_use.land_use_id IS 'Primary key for the land use types table.';


--
-- Name: COLUMN land_use.type; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.land_use.type IS 'The type of land use or resource being discussed (e.g., ''Fishing rights'').';


--
-- Name: land_use_land_use_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.land_use_land_use_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.land_use_land_use_id_seq OWNER TO gudrun;

--
-- Name: land_use_land_use_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.land_use_land_use_id_seq OWNED BY digidiggie.land_use.land_use_id;


--
-- Name: legal_sources; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.legal_sources (
    legal_source_id integer NOT NULL,
    legal_source_name text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie.legal_sources OWNER TO gudrun;

--
-- Name: COLUMN legal_sources.legal_source_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.legal_sources.legal_source_id IS 'Primary key for the legal sources table.';


--
-- Name: COLUMN legal_sources.legal_source_name; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.legal_sources.legal_source_name IS 'The name of the legal source or legal precedent cited.';


--
-- Name: legal_sources_legal_source_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.legal_sources_legal_source_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.legal_sources_legal_source_id_seq OWNER TO gudrun;

--
-- Name: legal_sources_legal_source_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.legal_sources_legal_source_id_seq OWNED BY digidiggie.legal_sources.legal_source_id;


--
-- Name: parishes; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.parishes (
    parish_id integer NOT NULL,
    parish text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie.parishes OWNER TO gudrun;

--
-- Name: COLUMN parishes.parish_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.parishes.parish_id IS 'Primary key for the parishes table.';


--
-- Name: COLUMN parishes.parish; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.parishes.parish IS 'Heading comes from the National Survey database of place names and means "parish or town", but only parishes are relevant here';


--
-- Name: parishes_parish_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.parishes_parish_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.parishes_parish_id_seq OWNER TO gudrun;

--
-- Name: parishes_parish_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.parishes_parish_id_seq OWNED BY digidiggie.parishes.parish_id;


--
-- Name: person_properties; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.person_properties (
    person_property_id integer NOT NULL,
    person_id integer,
    property_id integer,
    specifier text,
    property_value text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie.person_properties OWNER TO gudrun;

--
-- Name: person_properties_person_property_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.person_properties_person_property_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.person_properties_person_property_id_seq OWNER TO gudrun;

--
-- Name: person_properties_person_property_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.person_properties_person_property_id_seq OWNED BY digidiggie.person_properties.person_property_id;


--
-- Name: persons; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.persons (
    person_id integer NOT NULL,
    given_name text,
    patronymic text,
    surname text,
    birth_year integer,
    death_year integer,
    community_name text,
    full_name text GENERATED ALWAYS AS (((((COALESCE(given_name, ''::text) || ' '::text) || COALESCE(patronymic, ''::text)) || ' '::text) || COALESCE(surname, ''::text))) STORED
);


ALTER TABLE digidiggie.persons OWNER TO gudrun;

--
-- Name: COLUMN persons.person_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.persons.person_id IS 'Primary key for the persons table.';


--
-- Name: COLUMN persons.given_name; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.persons.given_name IS 'The person''s given name(s) (first name).';


--
-- Name: COLUMN persons.patronymic; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.persons.patronymic IS 'The person''s patronymic name (e.g., ''Andersson'').';


--
-- Name: COLUMN persons.surname; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.persons.surname IS 'The person''s surname, byname, or family name.';


--
-- Name: COLUMN persons.birth_year; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.persons.birth_year IS 'The year of birth.';


--
-- Name: COLUMN persons.death_year; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.persons.death_year IS 'The year of death.';


--
-- Name: COLUMN persons.community_name; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.persons.community_name IS 'The name of the village where the person primarily resided.';


--
-- Name: COLUMN persons.full_name; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.persons.full_name IS 'The person''s full name, likely a computed or concatenated field.';


--
-- Name: persons_person_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.persons_person_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.persons_person_id_seq OWNER TO gudrun;

--
-- Name: persons_person_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.persons_person_id_seq OWNED BY digidiggie.persons.person_id;


--
-- Name: placenames; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.placenames (
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
    geom_point digidiggie.geometry(Point,3006)
);


ALTER TABLE digidiggie.placenames OWNER TO gudrun;

--
-- Name: properties; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.properties (
    property_id integer NOT NULL,
    property_name text,
    description text
);


ALTER TABLE digidiggie.properties OWNER TO gudrun;

--
-- Name: properties_property_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.properties_property_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.properties_property_id_seq OWNER TO gudrun;

--
-- Name: properties_property_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.properties_property_id_seq OWNED BY digidiggie.properties.property_id;


--
-- Name: seasons; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.seasons (
    season_id integer NOT NULL,
    season_name text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie.seasons OWNER TO gudrun;

--
-- Name: COLUMN seasons.season_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.seasons.season_id IS 'Primary key for the seasons table.';


--
-- Name: COLUMN seasons.season_name; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.seasons.season_name IS 'The name of the season when the disputed resource was primarily used (e.g., ''Winter'', ''Summer'').';


--
-- Name: seasons_season_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.seasons_season_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.seasons_season_id_seq OWNER TO gudrun;

--
-- Name: seasons_season_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.seasons_season_id_seq OWNED BY digidiggie.seasons.season_id;


--
-- Name: sources; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.sources (
    source_id integer NOT NULL,
    source_name text DEFAULT ''::text NOT NULL,
    source_abbreviation character varying(255)
);


ALTER TABLE digidiggie.sources OWNER TO gudrun;

--
-- Name: COLUMN sources.source_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.sources.source_id IS 'Primary key for the sources table.';


--
-- Name: COLUMN sources.source_name; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.sources.source_name IS 'The full name of the historical source document.';


--
-- Name: COLUMN sources.source_abbreviation; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.sources.source_abbreviation IS 'The abbreviation for the source, used for quick reference.';


--
-- Name: sources_source_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.sources_source_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.sources_source_id_seq OWNER TO gudrun;

--
-- Name: sources_source_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.sources_source_id_seq OWNED BY digidiggie.sources.source_id;


--
-- Name: winners; Type: TABLE; Schema: public; Owner: gudrun
--

CREATE TABLE digidiggie.winners (
    winner_id integer NOT NULL,
    winner_description text DEFAULT ''::text NOT NULL
);


ALTER TABLE digidiggie.winners OWNER TO gudrun;

--
-- Name: COLUMN winners.winner_id; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.winners.winner_id IS 'Primary key for the winners table.';


--
-- Name: COLUMN winners.winner_description; Type: COMMENT; Schema: public; Owner: gudrun
--

COMMENT ON COLUMN digidiggie.winners.winner_description IS 'Describes the winning party in a legal case (e.g., ''settler'' or ''reindeer herder'')';


--
-- Name: winners_winner_id_seq; Type: SEQUENCE; Schema: public; Owner: gudrun
--

CREATE SEQUENCE digidiggie.winners_winner_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE digidiggie.winners_winner_id_seq OWNER TO gudrun;

--
-- Name: winners_winner_id_seq; Type: SEQUENCE OWNED BY; Schema: public; Owner: gudrun
--

ALTER SEQUENCE digidiggie.winners_winner_id_seq OWNED BY digidiggie.winners.winner_id;


--
-- Name: communities community_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.communities ALTER COLUMN community_id SET DEFAULT nextval('digidiggie.communities_community_id_seq'::regclass);


--
-- Name: entries entry_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries ALTER COLUMN entry_id SET DEFAULT nextval('digidiggie.entries_entry_id_seq'::regclass);


--
-- Name: judgements judgement_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.judgements ALTER COLUMN judgement_id SET DEFAULT nextval('digidiggie.judgements_judgement_id_seq'::regclass);


--
-- Name: land_use land_use_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.land_use ALTER COLUMN land_use_id SET DEFAULT nextval('digidiggie.land_use_land_use_id_seq'::regclass);


--
-- Name: legal_sources legal_source_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.legal_sources ALTER COLUMN legal_source_id SET DEFAULT nextval('digidiggie.legal_sources_legal_source_id_seq'::regclass);


--
-- Name: parishes parish_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.parishes ALTER COLUMN parish_id SET DEFAULT nextval('digidiggie.parishes_parish_id_seq'::regclass);


--
-- Name: person_properties person_property_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.person_properties ALTER COLUMN person_property_id SET DEFAULT nextval('digidiggie.person_properties_person_property_id_seq'::regclass);


--
-- Name: persons person_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.persons ALTER COLUMN person_id SET DEFAULT nextval('digidiggie.persons_person_id_seq'::regclass);


--
-- Name: properties property_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.properties ALTER COLUMN property_id SET DEFAULT nextval('digidiggie.properties_property_id_seq'::regclass);


--
-- Name: seasons season_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.seasons ALTER COLUMN season_id SET DEFAULT nextval('digidiggie.seasons_season_id_seq'::regclass);


--
-- Name: sources source_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.sources ALTER COLUMN source_id SET DEFAULT nextval('digidiggie.sources_source_id_seq'::regclass);


--
-- Name: winners winner_id; Type: DEFAULT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.winners ALTER COLUMN winner_id SET DEFAULT nextval('digidiggie.winners_winner_id_seq'::regclass);


--
-- Name: communities communities_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.communities
    ADD CONSTRAINT communities_pkey PRIMARY KEY (community_id);


--
-- Name: entries entries_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries
    ADD CONSTRAINT entries_pkey PRIMARY KEY (entry_id);


--
-- Name: judgements judgements_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.judgements
    ADD CONSTRAINT judgements_pkey PRIMARY KEY (judgement_id);


--
-- Name: land_use land_use_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.land_use
    ADD CONSTRAINT land_use_pkey PRIMARY KEY (land_use_id);


--
-- Name: legal_sources legal_sources_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.legal_sources
    ADD CONSTRAINT legal_sources_pkey PRIMARY KEY (legal_source_id);


--
-- Name: parishes parishes_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.parishes
    ADD CONSTRAINT parishes_pkey PRIMARY KEY (parish_id);


--
-- Name: person_properties person_properties_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.person_properties
    ADD CONSTRAINT person_properties_pkey PRIMARY KEY (person_property_id);


--
-- Name: persons persons_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.persons
    ADD CONSTRAINT persons_pkey PRIMARY KEY (person_id);


--
-- Name: placenames placenames_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.placenames
    ADD CONSTRAINT placenames_pkey PRIMARY KEY (fid);


--
-- Name: properties properties_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.properties
    ADD CONSTRAINT properties_pkey PRIMARY KEY (property_id);


--
-- Name: seasons seasons_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.seasons
    ADD CONSTRAINT seasons_pkey PRIMARY KEY (season_id);


--
-- Name: sources sources_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.sources
    ADD CONSTRAINT sources_pkey PRIMARY KEY (source_id);


--
-- Name: winners winners_pkey; Type: CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.winners
    ADD CONSTRAINT winners_pkey PRIMARY KEY (winner_id);


--
-- Name: communities_parish_id_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX communities_parish_id_idx ON digidiggie.communities USING btree (parish_id);


--
-- Name: entries_actor_id_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX entries_actor_id_idx ON digidiggie.entries USING btree (actor_id);


--
-- Name: entries_community_id_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX entries_community_id_idx ON digidiggie.entries USING btree (community_id);


--
-- Name: entries_judgement_id_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX entries_judgement_id_idx ON digidiggie.entries USING btree (judgement_id);


--
-- Name: entries_land_use_id_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX entries_land_use_id_idx ON digidiggie.entries USING btree (land_use_id);


--
-- Name: entries_legal_source_id_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX entries_legal_source_id_idx ON digidiggie.entries USING btree (legal_source_id);


--
-- Name: entries_placename_id_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX entries_placename_id_idx ON digidiggie.entries USING btree (placename_id);


--
-- Name: entries_reference_number_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX entries_reference_number_idx ON digidiggie.entries USING btree (reference_number);


--
-- Name: entries_season_id_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX entries_season_id_idx ON digidiggie.entries USING btree (season_id);


--
-- Name: entries_winner_id_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX entries_winner_id_idx ON digidiggie.entries USING btree (winner_id);


--
-- Name: parishes_parish_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX parishes_parish_idx ON digidiggie.parishes USING btree (parish);


--
-- Name: seasons_season_name_idx; Type: INDEX; Schema: public; Owner: gudrun
--

CREATE INDEX seasons_season_name_idx ON digidiggie.seasons USING btree (season_name);


--
-- Name: communities communities_parish_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.communities
    ADD CONSTRAINT communities_parish_id_fkey FOREIGN KEY (parish_id) REFERENCES digidiggie.parishes(parish_id);


--
-- Name: entries entries_actor_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries
    ADD CONSTRAINT entries_actor_id_fkey FOREIGN KEY (actor_id) REFERENCES digidiggie.persons(person_id);


--
-- Name: entries entries_community_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries
    ADD CONSTRAINT entries_community_id_fkey FOREIGN KEY (community_id) REFERENCES digidiggie.communities(community_id);


--
-- Name: entries entries_judgement_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries
    ADD CONSTRAINT entries_judgement_id_fkey FOREIGN KEY (judgement_id) REFERENCES digidiggie.judgements(judgement_id);


--
-- Name: entries entries_land_use_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries
    ADD CONSTRAINT entries_land_use_id_fkey FOREIGN KEY (land_use_id) REFERENCES digidiggie.land_use(land_use_id);


--
-- Name: entries entries_legal_source_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries
    ADD CONSTRAINT entries_legal_source_id_fkey FOREIGN KEY (legal_source_id) REFERENCES digidiggie.legal_sources(legal_source_id);


--
-- Name: entries entries_placename_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries
    ADD CONSTRAINT entries_placename_id_fkey FOREIGN KEY (placename_id) REFERENCES digidiggie.placenames(fid);


--
-- Name: entries entries_season_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries
    ADD CONSTRAINT entries_season_id_fkey FOREIGN KEY (season_id) REFERENCES digidiggie.seasons(season_id);


--
-- Name: entries entries_source_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries
    ADD CONSTRAINT entries_source_id_fkey FOREIGN KEY (source_id) REFERENCES digidiggie.sources(source_id);


--
-- Name: entries entries_winner_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.entries
    ADD CONSTRAINT entries_winner_id_fkey FOREIGN KEY (winner_id) REFERENCES digidiggie.winners(winner_id);


--
-- Name: person_properties person_properties_person_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.person_properties
    ADD CONSTRAINT person_properties_person_id_fkey FOREIGN KEY (person_id) REFERENCES digidiggie.persons(person_id);


--
-- Name: person_properties person_properties_property_id_fkey; Type: FK CONSTRAINT; Schema: public; Owner: gudrun
--

ALTER TABLE ONLY digidiggie.person_properties
    ADD CONSTRAINT person_properties_property_id_fkey FOREIGN KEY (property_id) REFERENCES digidiggie.properties(property_id);

