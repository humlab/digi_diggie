-- =====================================================================================================================
-- DigiDiggie TNG Schema - MS Access Version
-- =====================================================================================================================
-- 
-- This is an MS Access (Jet/ACE SQL) compatible version of the digidiggie_en_tng PostgreSQL schema.
-- 
-- IMPORTANT NOTES:
-- 1. MS Access does not support schemas - all tables are in the default namespace
-- 2. MS Access does not support PostGIS geometry types - geom column is omitted
-- 3. MS Access does not support COMMENT ON - comments are included as SQL comments
-- 4. MS Access does not support generated columns - full_name is a computed field in forms instead
-- 5. SERIAL data type is replaced with AUTOINCREMENT (or COUNTER in older versions)
-- 6. Some indexes and foreign key constraints may need to be created via Access UI
-- 
-- =====================================================================================================================

-- =====================================================================================================================
-- TABLE DEFINITIONS
-- =====================================================================================================================

-- Community table
-- Social or administrative grouping (e.g., parish or community) to which a person belongs at the time of the court case


CREATE TABLE community (
    community_id COUNTER PRIMARY KEY,           -- Primary key for the community table
    community_name TEXT NOT NULL,               -- The name of the community (village)
    parish_id LONG                              -- Foreign key linking to the parish this community belongs to
);

-- Court case table
-- Single court proceeding recorded in a historical source, identified by date, court, and source text. 
-- A court case results in exactly one ruling in this model
CREATE TABLE court_case (
    court_case_id COUNTER PRIMARY KEY,          -- Primary key for the court case table
    source_id LONG NOT NULL,                    -- Foreign key to the source document (e.g., court record collection)
    reference_number VARCHAR(16),               -- Reference number within a specific source collection, e.g., K.B. Wiklund's transcripts
    district_court_name TEXT,                   -- Name of the district court (tingslag) where the case was heard
    case_year LONG,                             -- The year the case was heard at court
    source_text LONGTEXT                        -- Aggregated description/text from the source document for this case
);

-- Court case entry table
-- Discrete unit of information extracted from a court case, typically describing a specific land-use situation, event, or claim
CREATE TABLE court_case_entry (
    court_case_entry_id COUNTER PRIMARY KEY,    -- Primary key for the court case entry table
    court_case_id LONG NOT NULL,                -- Foreign key linking to the parent court case
    entry_year LONG,                            -- The year the event occurred, or if unknown, the year the matter was heard at court
    curated_text LONGTEXT,                      -- Curated description of the event in free text
    original_placename TEXT,                    -- The place's name as written in the original source document
    season_id LONG,                             -- Foreign key to the season when the disputed resource was used
    land_use_id LONG,                           -- Foreign key to the resource or land use type involved in the dispute
    placename_id LONG                           -- Foreign key to a standardized placename from the Swedish National Survey (enables GIS connection)
);

-- Person table
-- Historical individual identified in the sources, with personal attributes where known (name, patronymic, birth/death year, notes)
-- NOTE: full_name is NOT a generated column in Access - use a query or form expression: [given_name] & " " & [patronymic] & " " & [surname]
CREATE TABLE person (
    person_id COUNTER PRIMARY KEY,              -- Primary key for the person table
    given_name TEXT,                            -- The person's given name(s) (first name)
    patronymic TEXT,                            -- The person's patronymic name (e.g., 'Andersson', 'Jonsdotter')
    surname TEXT,                               -- The person's surname, byname, or family name
    birth_year LONG,                            -- The year of birth
    death_year LONG,                            -- The year of death
    community_name TEXT,                        -- The name of the village where the person primarily resided
    note LONGTEXT                               -- Additional notes about the person
);

-- Land rights status table
-- Description of a person's legal or customary status to the land as interpreted from the entry (e.g., owned land, no land rights, uncertain)
CREATE TABLE land_rights_status (
    land_rights_status_id COUNTER PRIMARY KEY,  -- Primary key for the land rights status table
    land_rights_status VARCHAR(255) NOT NULL,   -- Land rights status name (e.g., 'Ja', 'Nej', 'Nja')
    description TEXT NOT NULL                   -- Description of the land rights status
);

-- Legal source table
-- Normative legal text (law code, regulation, precedent) that a ruling cites or applies
CREATE TABLE legal_source (
    legal_source_id COUNTER PRIMARY KEY,        -- Primary key for the legal source table
    legal_source_name TEXT NOT NULL             -- The name of the legal source or legal precedent cited
);

-- Land use table
-- Categorized description of how land is used or claimed in an entry (e.g., fishing, hunting, herding, reindeer grazing), 
-- based on interpretation of the source
CREATE TABLE land_use (
    land_use_id COUNTER PRIMARY KEY,            -- Primary key for the land use table
    description TEXT NOT NULL                   -- The type of land use or resource (e.g., 'Fishing rights', 'Reindeer grazing')
);

-- Parish table
-- Lookup table for parishes
CREATE TABLE parish (
    parish_id COUNTER PRIMARY KEY,              -- Primary key for the parish table
    parish TEXT NOT NULL                        -- Parish name - heading comes from the National Survey database of place names
);

-- Outcome type table
-- Categorization describing the kind of decision outcome (e.g., winner, sanction, injunction with fine, partition of land)
CREATE TABLE outcome_type (
    outcome_type_id COUNTER PRIMARY KEY,        -- Primary key for the outcome type table
    outcome_type_name TEXT NOT NULL UNIQUE,     -- Outcome type name (e.g., 'Vinnare', 'Böter', 'Friad')
    description TEXT                            -- Description of the outcome type
);

-- Person outcome table
-- Outcome of a ruling as it affects a specific person (e.g., being sanctioned, declared winner). Connects rulings to individuals
CREATE TABLE person_outcome (
    person_outcome_id COUNTER PRIMARY KEY,      -- Primary key for the person outcome table
    ruling_id LONG NOT NULL,                    -- Foreign key linking to the ruling
    person_id LONG NOT NULL,                    -- Foreign key linking to the person
    outcome_type_id LONG NOT NULL,              -- Foreign key to the outcome type (e.g., damages, fined, acquitted, winner)
    description TEXT                            -- Description of the specific outcome for this person
);

-- Placename table
-- Standardized geographical place associated with an entry, linked to the Swedish National Survey 
-- (external authority/placename registry) with coordinates
-- NOTE: PostGIS geometry column omitted - use separate coordinate columns for mapping
CREATE TABLE placename (
    placename_id COUNTER PRIMARY KEY,           -- Primary key for the placename table (id in original)
    placename TEXT,                             -- The standardized placename (ortnamn)
    northing LONG,                              -- Northing coordinate in SWEREF 99 TM (EPSG:3006) (n in original)
    easting LONG,                               -- Easting coordinate in SWEREF 99 TM (EPSG:3006) (e in original)
    serial_number TEXT,                         -- Serial number (löpnummer) from the National Survey (lopnr in original)
    name_type_code TEXT,                        -- Name type code (namntyp) from the National Survey (namntyp_nr in original)
    language_code TEXT,                         -- Language code (språk) from the National Survey (språk_nr in original)
    parish_code TEXT,                           -- Parish/town code (sockenstad) from the National Survey (sockenstad_nr in original)
    county_code TEXT,                           -- County code (län) from the National Survey (lan_nr in original)
    municipality_code TEXT,                     -- Municipality code (kommun) from the National Survey (kommun_nr in original)
    combined_placename TEXT,                    -- Combined placename (combination of multiple name components) (kombo in original)
    parish_name TEXT                            -- Parish name from the National Survey
    -- geom geometry(Point, 4326) omitted - not supported in MS Access
    -- Use latitude/longitude columns if needed for WGS84 coordinates (EPSG:4326)
);

-- Person entry table
-- Contextualized appearance of a person within a specific court case entry. Captures the person's role, 
-- land rights status, and how they are described in the source
CREATE TABLE person_entry (
    person_entry_id COUNTER PRIMARY KEY,        -- Primary key for the person entry table
    court_case_entry_id LONG NOT NULL,          -- Foreign key linking to the court case entry
    person_id LONG NOT NULL,                    -- Foreign key linking to the person involved
    community_id LONG,                          -- Foreign key to the community where the person resided at the time
    land_rights_status_id LONG NOT NULL,        -- Foreign key indicating if the person had land rights
    role_id LONG,                               -- Foreign key to the person's role in the case (e.g., plaintiff, defendant, witness)
    curated_text LONGTEXT                       -- Additional curated text describing the person's involvement
);

-- Role type table
-- Lookup table for role type categories (social or judicial)
CREATE TABLE role_type (
    role_type_id COUNTER PRIMARY KEY,           -- Primary key for the role type table
    role_type_name TEXT NOT NULL UNIQUE,        -- Role type category name ('Social' or 'Juridisk')
    description TEXT NOT NULL                   -- Description of the role type category
);

-- Role table
-- Social or legal role attributed to a person in a specific entry (e.g., Nybyggare, Sámi, plaintiff, defendant). 
-- Roles are contextual, not permanent identities
CREATE TABLE role (
    role_id COUNTER PRIMARY KEY,                -- Primary key for the role table
    role_name TEXT NOT NULL UNIQUE,             -- Role name (e.g., 'Klagande', 'Svarande', 'Vittne', 'Same', 'Nybyggare')
    role_type_id LONG NOT NULL,                 -- Foreign key to role type (social or judicial)
    description TEXT NOT NULL                   -- Description of the role
);

-- Season table
-- Optional temporal qualifier indicating when the event described in an entry took place (e.g., summer, winter)
CREATE TABLE season (
    season_id COUNTER PRIMARY KEY,              -- Primary key for the season table
    season_name TEXT NOT NULL UNIQUE            -- The name of the season when the disputed resource was primarily used (e.g., 'Winter', 'Summer')
);

-- Ruling type table
-- Lookup table for ruling/judgement types
CREATE TABLE ruling_type (
    ruling_type_id COUNTER PRIMARY KEY,         -- Primary key for the ruling type table
    ruling_type VARCHAR(255) NOT NULL UNIQUE,   -- Ruling type name (e.g., 'Dom', 'Förlikning', 'Hänvisning')
    description TEXT NOT NULL                   -- Description of the ruling type
);

-- Ruling table
-- Judicial decision resulting from a court case. Records the year, description, ruling type (resolved or referred), 
-- and may cite a legal source
CREATE TABLE ruling (
    ruling_id COUNTER PRIMARY KEY,              -- Primary key for the ruling table
    court_case_id LONG NOT NULL UNIQUE,         -- Foreign key linking to the court case (one-to-one relationship)
    ruling_year LONG,                           -- The year the ruling was issued
    description TEXT,                           -- Description of the ruling in free text
    ruling_type_id LONG NOT NULL,               -- Foreign key to the type of ruling (e.g., judgement, settlement, referral)
    legal_source_id LONG                        -- Foreign key to the legal source or precedent cited in the ruling
);

-- Source table
-- Historical source from which court cases are excerpted (e.g., court records, archival volumes), with identifiers and metadata
CREATE TABLE source (
    source_id COUNTER PRIMARY KEY,              -- Primary key for the source table
    source_name TEXT NOT NULL,                  -- The full name of the historical source document
    source_abbreviation VARCHAR(255)            -- The abbreviation for the source, used for quick reference (e.g., 'DB' for court record)
);

-- Person relationship table
-- Relationships between persons (family, social connections)
CREATE TABLE person_relationship (
    person_relationship_id COUNTER PRIMARY KEY, -- Primary key for the person relationship table
    person_1_id LONG NOT NULL,                  -- Foreign key to the first person in the relationship
    person_2_id LONG NOT NULL,                  -- Foreign key to the second person in the relationship
    relationship_type_id LONG NOT NULL,         -- Foreign key to the type of relationship (e.g., father, mother, sibling, spouse)
    description TEXT                            -- Additional description of the relationship
);

-- Relationship type table
-- Lookup table for person relationship types
CREATE TABLE relationship_type (
    relationship_type_id COUNTER PRIMARY KEY,   -- Primary key for the relationship type table
    relationship_type_name TEXT NOT NULL UNIQUE,-- Relationship type name (e.g., 'Far', 'Mor', 'Syskon', 'Make/Maka')
    description TEXT NOT NULL                   -- Description of the relationship type
);

-- =====================================================================================================================
-- INDEXES
-- =====================================================================================================================
-- NOTE: These should be created via Access UI or using CREATE INDEX syntax
-- Some indexes on UNIQUE columns may already exist from PRIMARY KEY or UNIQUE constraints

CREATE INDEX communities_parish_id_idx ON community (parish_id);
CREATE INDEX court_cases_case_year_idx ON court_case (case_year);
CREATE INDEX court_cases_reference_number_idx ON court_case (reference_number);
CREATE INDEX court_cases_source_id_idx ON court_case (source_id);
CREATE INDEX entries_court_case_id_idx ON court_case_entry (court_case_id);
CREATE INDEX entries_land_use_id_idx ON court_case_entry (land_use_id);
CREATE INDEX entries_placename_id_idx ON court_case_entry (placename_id);
CREATE INDEX entries_season_id_idx ON court_case_entry (season_id);
CREATE INDEX rulings_legal_source_id_idx ON ruling (legal_source_id);
CREATE INDEX parishes_parish_idx ON parish (parish);
CREATE INDEX person_entries_person_id_idx ON person_entry (person_id);
CREATE INDEX person_entries_community_id_idx ON person_entry (community_id);
CREATE INDEX person_entries_court_case_entry_id_idx ON person_entry (court_case_entry_id);
CREATE INDEX person_outcomes_outcome_type_id_idx ON person_outcome (outcome_type_id);
CREATE INDEX person_outcomes_person_id_idx ON person_outcome (person_id);
CREATE INDEX person_outcomes_ruling_id_idx ON person_outcome (ruling_id);
CREATE INDEX seasons_season_name_idx ON season (season_name);

-- =====================================================================================================================
-- FOREIGN KEY CONSTRAINTS
-- =====================================================================================================================
-- NOTE: Foreign key constraints in Access are typically created via Relationships window in Access UI
-- The syntax below may work in some versions of Access but is not universally supported
-- If these fail, create relationships manually in Access via Database Tools > Relationships

-- ALTER TABLE community ADD CONSTRAINT community_parish_id_fkey FOREIGN KEY (parish_id) REFERENCES parish (parish_id);
-- ALTER TABLE court_case ADD CONSTRAINT court_case_source_id_fkey FOREIGN KEY (source_id) REFERENCES source (source_id);
-- ALTER TABLE court_case_entry ADD CONSTRAINT entry_court_case_id_fkey FOREIGN KEY (court_case_id) REFERENCES court_case (court_case_id);
-- ALTER TABLE court_case_entry ADD CONSTRAINT entry_land_use_id_fkey FOREIGN KEY (land_use_id) REFERENCES land_use (land_use_id);
-- ALTER TABLE court_case_entry ADD CONSTRAINT entry_season_id_fkey FOREIGN KEY (season_id) REFERENCES season (season_id);
-- ALTER TABLE court_case_entry ADD CONSTRAINT entry_placename_id_fkey FOREIGN KEY (placename_id) REFERENCES placename (placename_id);
-- ALTER TABLE person_entry ADD CONSTRAINT person_entry_person_id_fkey FOREIGN KEY (person_id) REFERENCES person (person_id);
-- ALTER TABLE person_entry ADD CONSTRAINT person_entry_community_id_fkey FOREIGN KEY (community_id) REFERENCES community (community_id);
-- ALTER TABLE person_entry ADD CONSTRAINT person_entry_entry_id_fkey FOREIGN KEY (court_case_entry_id) REFERENCES court_case_entry (court_case_entry_id);
-- ALTER TABLE person_entry ADD CONSTRAINT person_entry_role_id_fkey FOREIGN KEY (role_id) REFERENCES role (role_id);
-- ALTER TABLE person_entry ADD CONSTRAINT fk_person_entry_land_rights_status_1 FOREIGN KEY (land_rights_status_id) REFERENCES land_rights_status (land_rights_status_id);
-- ALTER TABLE role ADD CONSTRAINT role_role_type_id_fkey FOREIGN KEY (role_type_id) REFERENCES role_type (role_type_id);
-- ALTER TABLE person_outcome ADD CONSTRAINT person_outcomes_outcome_type_id_fkey FOREIGN KEY (outcome_type_id) REFERENCES outcome_type (outcome_type_id);
-- ALTER TABLE person_outcome ADD CONSTRAINT person_outcomes_person_id_fkey FOREIGN KEY (person_id) REFERENCES person (person_id);
-- ALTER TABLE person_outcome ADD CONSTRAINT person_outcomes_ruling_id_fkey FOREIGN KEY (ruling_id) REFERENCES ruling (ruling_id);
-- ALTER TABLE ruling ADD CONSTRAINT rulings_court_case_id_fkey FOREIGN KEY (court_case_id) REFERENCES court_case (court_case_id);
-- ALTER TABLE ruling ADD CONSTRAINT rulings_legal_source_id_fkey FOREIGN KEY (legal_source_id) REFERENCES legal_source (legal_source_id);
-- ALTER TABLE ruling ADD CONSTRAINT fk_rulings_ruling_type_1 FOREIGN KEY (ruling_type_id) REFERENCES ruling_type (ruling_type_id);
-- ALTER TABLE person_relationship ADD CONSTRAINT person_relationship_person_1_id_fkey FOREIGN KEY (person_1_id) REFERENCES person (person_id);
-- ALTER TABLE person_relationship ADD CONSTRAINT person_relationship_person_2_id_fkey FOREIGN KEY (person_2_id) REFERENCES person (person_id);
-- ALTER TABLE person_relationship ADD CONSTRAINT person_relationship_relationship_type_id_fkey FOREIGN KEY (relationship_type_id) REFERENCES relationship_type (relationship_type_id);

-- =====================================================================================================================
-- ADDITIONAL CONSTRAINTS (UNIQUE)
-- =====================================================================================================================
-- NOTE: UNIQUE constraints may need to be created via Access UI or as separate indexes
-- Some UNIQUE constraints are already defined inline in the CREATE TABLE statements above

-- ALTER TABLE court_case ADD CONSTRAINT court_cases_source_id_key UNIQUE (source_id, reference_number, case_year);

-- =====================================================================================================================
-- NOTES ON MS ACCESS MIGRATION
-- =====================================================================================================================
--
-- 1. COUNTER vs AUTOINCREMENT:
--    - COUNTER is the Access/Jet SQL data type
--    - AUTOINCREMENT is the ANSI SQL-92 syntax (also supported in newer versions)
--    - Both create auto-incrementing integer primary keys
--
-- 2. TEXT vs LONGTEXT vs MEMO:
--    - TEXT (or VARCHAR) stores up to 255 characters
--    - LONGTEXT (or MEMO in older Access) stores up to 65,535 characters
--    - Use LONGTEXT for potentially long text fields (curated_text, description, notes, etc.)
--
-- 3. LONG vs INTEGER:
--    - LONG is Access equivalent to 32-bit INTEGER
--    - Use LONG for foreign keys and integer fields
--
-- 4. Generated Columns:
--    - MS Access does NOT support generated/computed columns at the table level
--    - Use calculated fields in queries or forms instead
--    - Example for person.full_name in a query:
--      SELECT *, [given_name] & " " & [patronymic] & " " & [surname] AS full_name FROM person
--
-- 5. Foreign Key Relationships:
--    - Best practice: Create relationships via Access UI (Database Tools > Relationships)
--    - Enforce referential integrity, cascade updates/deletes as needed
--    - Some versions of Access support ALTER TABLE ... ADD CONSTRAINT ... FOREIGN KEY syntax
--
-- 6. PostGIS Geometry:
--    - MS Access does NOT support PostGIS geometry types
--    - Store coordinates as separate numeric columns (latitude, longitude, northing, easting)
--    - Use external GIS tools or linked PostgreSQL tables for spatial queries
--
-- 7. Encoding:
--    - MS Access uses Windows-1252 or UTF-16 encoding internally
--    - Swedish characters (åäöÅÄÖ) should work without issues
--    - Be careful when importing/exporting to CSV (use UTF-8 with BOM)
--
-- 8. Schema/Namespace:
--    - MS Access does NOT support database schemas
--    - All tables are in the default namespace
--    - No "digidiggie_tng." prefix needed
--
-- 9. Functions and Procedures:
--    - MS Access does NOT support PostgreSQL-style user-defined functions
--    - Use VBA functions or queries for similar functionality
--    - export_schema_as_json() and sync_sequences() procedures omitted
--
-- 10. Default Values:
--     - MS Access supports DEFAULT values but syntax may differ
--     - Empty string defaults (''::text) are simplified to empty string or omitted
--
-- =====================================================================================================================
-- END OF SCHEMA
-- =====================================================================================================================
