-- =====================================================================================================================
-- DigiDiggie TNG - Migration Script with COUNTER Primary Keys
-- =====================================================================================================================
-- 
-- This script migrates existing Access tables while PRESERVING original primary key values
-- AND enabling COUNTER (auto-increment) for new records.
--
-- PROCESS FOR EACH TABLE:
-- 1. Create temp_tablename as a copy of tablename (with data)
-- 2. Drop original tablename
-- 3. Create tablename with COUNTER PRIMARY KEY
-- 4. Insert all data from temp_tablename INCLUDING original primary key values
-- 5. Drop temp_tablename
--
-- HOW IT WORKS:
-- - Access allows inserting explicit values into COUNTER fields
-- - After insertion, auto-increment continues from the highest inserted value
-- - Example: Insert IDs 1,2,5,10 → next auto-generated ID is 11
--
-- IMPORTANT NOTES:
-- - Original ID values from PostgreSQL are PRESERVED (FK relationships intact)
-- - New records added in Access will auto-increment from max(existing_id)+1
-- - After migration, re-create foreign key constraints via Access UI
--
-- WARNING: This is a destructive operation. Back up your database before running!
-- =====================================================================================================================

-- =====================================================================================================================
-- COMMUNITY TABLE
-- =====================================================================================================================

-- Step 1: Create temporary copy with data
SELECT * INTO temp_community FROM community;

-- Step 2: Drop original table
DROP TABLE community;

-- Step 3: Create new table with COUNTER primary key
CREATE TABLE community (
    community_id COUNTER PRIMARY KEY,
    community_name TEXT NOT NULL,
    parish_id LONG
);

-- Step 4: Insert data from temp table (including original IDs, counter continues from max)
INSERT INTO community (community_id, community_name, parish_id)
SELECT community_id, community_name, parish_id
FROM temp_community;

-- Step 5: Drop temporary table
DROP TABLE temp_community;

-- =====================================================================================================================
-- COURT_CASE TABLE
-- =====================================================================================================================

SELECT * INTO temp_court_case FROM court_case;

DROP TABLE court_case;

CREATE TABLE court_case (
    court_case_id COUNTER PRIMARY KEY,
    source_id LONG NOT NULL,
    reference_number VARCHAR(16),
    district_court_name TEXT,
    case_year LONG,
    source_text LONGTEXT
);

INSERT INTO court_case (court_case_id, source_id, reference_number, district_court_name, case_year, source_text)
SELECT court_case_id, source_id, reference_number, district_court_name, case_year, source_text
FROM temp_court_case;

DROP TABLE temp_court_case;

-- =====================================================================================================================
-- COURT_CASE_ENTRY TABLE
-- =====================================================================================================================

SELECT * INTO temp_court_case_entry FROM court_case_entry;

DROP TABLE court_case_entry;

CREATE TABLE court_case_entry (
    court_case_entry_id COUNTER PRIMARY KEY,
    court_case_id LONG NOT NULL,
    entry_year LONG,
    curated_text LONGTEXT,
    original_placename TEXT,
    season_id LONG,
    land_use_id LONG,
    placename_id LONG
);

INSERT INTO court_case_entry (court_case_entry_id, court_case_id, entry_year, curated_text, original_placename, season_id, land_use_id, placename_id)
SELECT court_case_entry_id, court_case_id, entry_year, curated_text, original_placename, season_id, land_use_id, placename_id
FROM temp_court_case_entry;

DROP TABLE temp_court_case_entry;

-- =====================================================================================================================
-- PERSON TABLE
-- =====================================================================================================================

SELECT * INTO temp_person FROM person;

DROP TABLE person;

CREATE TABLE person (
    person_id COUNTER PRIMARY KEY,
    given_name TEXT,
    patronymic TEXT,
    surname TEXT,
    birth_year LONG,
    death_year LONG,
    community_name TEXT,
    note LONGTEXT
);

INSERT INTO person (person_id, given_name, patronymic, surname, birth_year, death_year, community_name, note)
SELECT person_id, given_name, patronymic, surname, birth_year, death_year, community_name, note
FROM temp_person;

DROP TABLE temp_person;

-- =====================================================================================================================
-- LAND_RIGHTS_STATUS TABLE
-- =====================================================================================================================

SELECT * INTO temp_land_rights_status FROM land_rights_status;

DROP TABLE land_rights_status;

CREATE TABLE land_rights_status (
    land_rights_status_id COUNTER PRIMARY KEY,
    land_rights_status VARCHAR(255) NOT NULL,
    description TEXT NOT NULL
);

INSERT INTO land_rights_status (land_rights_status_id, land_rights_status, description)
SELECT land_rights_status_id, land_rights_status, description
FROM temp_land_rights_status;

DROP TABLE temp_land_rights_status;

-- =====================================================================================================================
-- LEGAL_SOURCE TABLE
-- =====================================================================================================================

SELECT * INTO temp_legal_source FROM legal_source;

DROP TABLE legal_source;

CREATE TABLE legal_source (
    legal_source_id COUNTER PRIMARY KEY,
    legal_source_name TEXT NOT NULL
);

INSERT INTO legal_source (legal_source_id, legal_source_name)
SELECT legal_source_id, legal_source_name
FROM temp_legal_source;

DROP TABLE temp_legal_source;

-- =====================================================================================================================
-- LAND_USE TABLE
-- =====================================================================================================================

SELECT * INTO temp_land_use FROM land_use;

DROP TABLE land_use;

CREATE TABLE land_use (
    land_use_id COUNTER PRIMARY KEY,
    description TEXT NOT NULL
);

INSERT INTO land_use (land_use_id, description)
SELECT land_use_id, description
FROM temp_land_use;

DROP TABLE temp_land_use;

-- =====================================================================================================================
-- PARISH TABLE
-- =====================================================================================================================

SELECT * INTO temp_parish FROM parish;

DROP TABLE parish;

CREATE TABLE parish (
    parish_id COUNTER PRIMARY KEY,
    parish TEXT NOT NULL
);

INSERT INTO parish (parish_id, parish)
SELECT parish_id, parish
FROM temp_parish;

DROP TABLE temp_parish;

-- =====================================================================================================================
-- OUTCOME_TYPE TABLE
-- =====================================================================================================================

SELECT * INTO temp_outcome_type FROM outcome_type;

DROP TABLE outcome_type;

CREATE TABLE outcome_type (
    outcome_type_id COUNTER PRIMARY KEY,
    outcome_type_name TEXT NOT NULL UNIQUE,
    description TEXT
);

INSERT INTO outcome_type (outcome_type_id, outcome_type_name, description)
SELECT outcome_type_id, outcome_type_name, description
FROM temp_outcome_type;

DROP TABLE temp_outcome_type;

-- =====================================================================================================================
-- PERSON_OUTCOME TABLE
-- =====================================================================================================================

SELECT * INTO temp_person_outcome FROM person_outcome;

DROP TABLE person_outcome;

CREATE TABLE person_outcome (
    person_outcome_id COUNTER PRIMARY KEY,
    ruling_id LONG NOT NULL,
    person_id LONG NOT NULL,
    outcome_type_id LONG NOT NULL,
    description TEXT
);

INSERT INTO person_outcome (person_outcome_id, ruling_id, person_id, outcome_type_id, description)
SELECT person_outcome_id, ruling_id, person_id, outcome_type_id, description
FROM temp_person_outcome;

DROP TABLE temp_person_outcome;

-- =====================================================================================================================
-- PLACENAME TABLE
-- =====================================================================================================================

SELECT * INTO temp_placename FROM placename;

DROP TABLE placename;

CREATE TABLE placename (
    placename_id COUNTER PRIMARY KEY,
    placename TEXT,
    northing LONG,
    easting LONG,
    serial_number TEXT,
    name_type_code TEXT,
    language_code TEXT,
    parish_code TEXT,
    county_code TEXT,
    municipality_code TEXT,
    combined_placename TEXT,
    parish_name TEXT
);

INSERT INTO placename (placename_id, placename, northing, easting, serial_number, name_type_code, language_code, parish_code, county_code, municipality_code, combined_placename, parish_name)
SELECT placename_id, placename, northing, easting, serial_number, name_type_code, language_code, parish_code, county_code, municipality_code, combined_placename, parish_name
FROM temp_placename;

DROP TABLE temp_placename;

-- =====================================================================================================================
-- PERSON_ENTRY TABLE
-- =====================================================================================================================

SELECT * INTO temp_person_entry FROM person_entry;

DROP TABLE person_entry;

CREATE TABLE person_entry (
    person_entry_id COUNTER PRIMARY KEY,
    court_case_entry_id LONG NOT NULL,
    person_id LONG NOT NULL,
    community_id LONG,
    land_rights_status_id LONG NOT NULL,
    role_id LONG,
    curated_text LONGTEXT
);

INSERT INTO person_entry (person_entry_id, court_case_entry_id, person_id, community_id, land_rights_status_id, role_id, curated_text)
SELECT person_entry_id, court_case_entry_id, person_id, community_id, land_rights_status_id, role_id, curated_text
FROM temp_person_entry;

DROP TABLE temp_person_entry;

-- =====================================================================================================================
-- ROLE_TYPE TABLE
-- =====================================================================================================================

SELECT * INTO temp_role_type FROM role_type;

DROP TABLE role_type;

CREATE TABLE role_type (
    role_type_id COUNTER PRIMARY KEY,
    role_type_name TEXT NOT NULL UNIQUE,
    description TEXT NOT NULL
);

INSERT INTO role_type (role_type_id, role_type_name, description)
SELECT role_type_id, role_type_name, description
FROM temp_role_type;

DROP TABLE temp_role_type;

-- =====================================================================================================================
-- ROLE TABLE
-- =====================================================================================================================

SELECT * INTO temp_role FROM role;

DROP TABLE role;

CREATE TABLE role (
    role_id COUNTER PRIMARY KEY,
    role_name TEXT NOT NULL UNIQUE,
    role_type_id LONG NOT NULL,
    description TEXT NOT NULL
);

INSERT INTO role (role_id, role_name, role_type_id, description)
SELECT role_id, role_name, role_type_id, description
FROM temp_role;

DROP TABLE temp_role;

-- =====================================================================================================================
-- SEASON TABLE
-- =====================================================================================================================

SELECT * INTO temp_season FROM season;

DROP TABLE season;

CREATE TABLE season (
    season_id COUNTER PRIMARY KEY,
    season_name TEXT NOT NULL UNIQUE
);

INSERT INTO season (season_id, season_name)
SELECT season_id, season_name
FROM temp_season;

DROP TABLE temp_season;

-- =====================================================================================================================
-- RULING_TYPE TABLE
-- =====================================================================================================================

SELECT * INTO temp_ruling_type FROM ruling_type;

DROP TABLE ruling_type;

CREATE TABLE ruling_type (
    ruling_type_id COUNTER PRIMARY KEY,
    ruling_type VARCHAR(255) NOT NULL UNIQUE,
    description TEXT NOT NULL
);

INSERT INTO ruling_type (ruling_type_id, ruling_type, description)
SELECT ruling_type_id, ruling_type, description
FROM temp_ruling_type;

DROP TABLE temp_ruling_type;

-- =====================================================================================================================
-- RULING TABLE
-- =====================================================================================================================

SELECT * INTO temp_ruling FROM ruling;

DROP TABLE ruling;

CREATE TABLE ruling (
    ruling_id COUNTER PRIMARY KEY,
    court_case_id LONG NOT NULL UNIQUE,
    ruling_year LONG,
    description TEXT,
    ruling_type_id LONG NOT NULL,
    legal_source_id LONG
);

INSERT INTO ruling (ruling_id, court_case_id, ruling_year, description, ruling_type_id, legal_source_id)
SELECT ruling_id, court_case_id, ruling_year, description, ruling_type_id, legal_source_id
FROM temp_ruling;

DROP TABLE temp_ruling;

-- =====================================================================================================================
-- SOURCE TABLE
-- =====================================================================================================================

SELECT * INTO temp_source FROM source;

DROP TABLE source;

CREATE TABLE source (
    source_id COUNTER PRIMARY KEY,
    source_name TEXT NOT NULL,
    source_abbreviation VARCHAR(255)
);

INSERT INTO source (source_id, source_name, source_abbreviation)
SELECT source_id, source_name, source_abbreviation
FROM temp_source;

DROP TABLE temp_source;

-- =====================================================================================================================
-- PERSON_RELATIONSHIP TABLE
-- =====================================================================================================================

SELECT * INTO temp_person_relationship FROM person_relationship;

DROP TABLE person_relationship;

CREATE TABLE person_relationship (
    person_relationship_id COUNTER PRIMARY KEY,
    person_1_id LONG NOT NULL,
    person_2_id LONG NOT NULL,
    relationship_type_id LONG NOT NULL,
    description TEXT
);

INSERT INTO person_relationship (person_relationship_id, person_1_id, person_2_id, relationship_type_id, description)
SELECT person_relationship_id, person_1_id, person_2_id, relationship_type_id, description
FROM temp_person_relationship;

DROP TABLE temp_person_relationship;

-- =====================================================================================================================
-- RELATIONSHIP_TYPE TABLE
-- =====================================================================================================================

SELECT * INTO temp_relationship_type FROM relationship_type;

DROP TABLE relationship_type;

CREATE TABLE relationship_type (
    relationship_type_id COUNTER PRIMARY KEY,
    relationship_type_name TEXT NOT NULL UNIQUE,
    description TEXT NOT NULL
);

INSERT INTO relationship_type (relationship_type_id, relationship_type_name, description)
SELECT relationship_type_id, relationship_type_name, description
FROM temp_relationship_type;

DROP TABLE temp_relationship_type;

-- =====================================================================================================================
-- MIGRATION COMPLETE
-- =====================================================================================================================
-- 
-- NEXT STEPS:
-- 1. Verify all records were migrated correctly by counting rows in each table
-- 2. Verify primary key values match the original data (critical for FK integrity)
-- 3. Re-create foreign key relationships via Access UI (Database Tools > Relationships)
-- 4. Re-create indexes if necessary
-- 5. Compact and repair the database (Database Tools > Compact and Repair)
--
-- WHAT YOU NOW HAVE:
-- ✓ Original ID values PRESERVED from PostgreSQL (foreign key relationships intact)
-- ✓ COUNTER (auto-increment) enabled for NEW records
-- ✓ New records will auto-increment starting from max(existing_id)+1
--
-- EXAMPLE: If a table has IDs 1,2,5,10 after migration, the next inserted record gets ID 11
--
-- This is the best of both worlds - preserved data integrity AND auto-increment functionality!
-- =====================================================================================================================
