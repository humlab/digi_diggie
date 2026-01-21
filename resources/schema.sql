-- ----------------------------------------------------------
-- MDB Tools - A library for reading MS Access database files
-- Copyright (C) 2000-2011 Brian Bruns and others.
-- Files in libmdb are licensed under LGPL and the utilities under
-- the GPL, see COPYING.LIB and COPYING files respectively.
-- Check out http://mdbtools.sourceforge.net
-- ----------------------------------------------------------

SET client_encoding = 'UTF-8';

DROP TABLE IF EXISTS "public"."seasons";
CREATE TABLE IF NOT EXISTS "public"."seasons"
 (
	"season_id"			SERIAL, 
	"season_name"			VARCHAR (255)
);

-- CREATE INDEXES ...
CREATE INDEX "seasons_årstid_idx" ON "public"."seasons" ("season_name");
ALTER TABLE "public"."seasons" ADD CONSTRAINT "seasons_pkey" PRIMARY KEY ("season_id");

DROP TABLE IF EXISTS "public"."communities";
CREATE TABLE IF NOT EXISTS "public"."communities"
 (
	"community_id"			SERIAL, 
	"community_name"			VARCHAR (255), 
	"parish_id"			INTEGER
);

-- CREATE INDEXES ...
ALTER TABLE "public"."communities" ADD CONSTRAINT "communities_pkey" PRIMARY KEY ("community_id");

DROP TABLE IF EXISTS "public"."entries";
CREATE TABLE IF NOT EXISTS "public"."entries"
 (
	"entry_id"			SERIAL, 
	"actor_id"			INTEGER, 
	"community_id"			INTEGER, 
	"year"			DOUBLE PRECISION, 
	"description"			VARCHAR (250), 
	"land_rights_status"			VARCHAR (255), 
	"source_id"			INTEGER, 
	"reference_number"			VARCHAR (16), 
	"season_id"			INTEGER, 
	"land_use_id"			INTEGER, 
	"original_placename"			VARCHAR (50), 
	"winner_id"			INTEGER, 
	"legal_source_id"			INTEGER, 
	"judgement_id"			INTEGER, 
	"placename_id"			INTEGER, 
	"lay_judge_involved"			BOOLEAN NOT NULL DEFAULT FALSE
);

-- CREATE INDEXES ...
ALTER TABLE "public"."entries" ADD CONSTRAINT "entries_pkey" PRIMARY KEY ("entry_id");
CREATE INDEX "entries_refnr_idx" ON "public"."entries" ("reference_number");

DROP TABLE IF EXISTS "public"."sources";
CREATE TABLE IF NOT EXISTS "public"."sources"
 (
	"source_id"			SERIAL, 
	"source_name"			VARCHAR (255), 
	"source_abbreviation"			VARCHAR (255)
);

-- CREATE INDEXES ...
ALTER TABLE "public"."sources" ADD CONSTRAINT "sources_pkey" PRIMARY KEY ("source_id");

DROP TABLE IF EXISTS "public"."land_use";
CREATE TABLE IF NOT EXISTS "public"."land_use"
 (
	"land_use_id"			SERIAL, 
	"type"			VARCHAR (255)
);

-- CREATE INDEXES ...
ALTER TABLE "public"."land_use" ADD CONSTRAINT "land_use_pkey" PRIMARY KEY ("land_use_id");

DROP TABLE IF EXISTS "public"."persons";
CREATE TABLE IF NOT EXISTS "public"."persons"
 (
	"person_id"			SERIAL, 
	"individual_id"			VARCHAR (255), 
	"father_id"			VARCHAR (255), 
	"mother_id"			VARCHAR (255), 
	"given_name"			VARCHAR (255), 
	"patronymic"			VARCHAR (255), 
	"surname"			VARCHAR (255), 
	"birth_date"			INTEGER, 
	"birth_year"			INTEGER, 
	"birth_place"			VARCHAR (255), 
	"residence_date"			INTEGER, 
	"death_date"			VARCHAR (255), 
	"death_year"			INTEGER, 
	"death_place"			VARCHAR (255), 
	"event_date"			VARCHAR (255), 
	"event_id"			INTEGER, 
	"community_name"			VARCHAR (255), 
	"full_name"			VARCHAR (254)
);

-- CREATE INDEXES ...
CREATE INDEX "persons_id_father_idx" ON "public"."persons" ("father_id");
CREATE INDEX "persons_id_individu_idx" ON "public"."persons" ("individual_id");
CREATE INDEX "persons_id_mother_idx" ON "public"."persons" ("mother_id");
ALTER TABLE "public"."persons" ADD CONSTRAINT "persons_pkey" PRIMARY KEY ("person_id");

DROP TABLE IF EXISTS "public"."legal_sources";
CREATE TABLE IF NOT EXISTS "public"."legal_sources"
 (
	"legal_source_id"			SERIAL, 
	"legal_source_name"			VARCHAR (255)
);

-- CREATE INDEXES ...
ALTER TABLE "public"."legal_sources" ADD CONSTRAINT "legal_sources_pkey" PRIMARY KEY ("legal_source_id");

DROP TABLE IF EXISTS "public"."parishes";
CREATE TABLE IF NOT EXISTS "public"."parishes"
 (
	"parish"			VARCHAR (255), 
	"parish_id"			SERIAL
);

-- CREATE INDEXES ...
CREATE INDEX "parishes_norm_orderbyindex_idx" ON "public"."parishes" ("parish");
ALTER TABLE "public"."parishes" ADD CONSTRAINT "parishes_pkey" PRIMARY KEY ("parish_id");

DROP TABLE IF EXISTS "public"."winners";
CREATE TABLE IF NOT EXISTS "public"."winners"
 (
	"winner_id"			SERIAL, 
	"winner_description"			VARCHAR (255)
);

-- CREATE INDEXES ...
ALTER TABLE "public"."winners" ADD CONSTRAINT "winners_pkey" PRIMARY KEY ("winner_id");

DROP TABLE IF EXISTS "public"."translationmapping";
CREATE TABLE IF NOT EXISTS "public"."translationmapping"
 (
	"originaltable"			VARCHAR (255), 
	"translatedtable"			VARCHAR (255), 
	"originalcolumn"			VARCHAR (255), 
	"translatedcolumn"			VARCHAR (255), 
	"comment"			VARCHAR (255), 
	"deprecateflag"			VARCHAR (255), 
	"translatedexpression"			VARCHAR (255)
);

-- CREATE INDEXES ...

DROP TABLE IF EXISTS "public"."rowsource";
CREATE TABLE IF NOT EXISTS "public"."rowsource"
 (
	"originaltable"			VARCHAR (255), 
	"originalcolumn"			VARCHAR (255), 
	"translatedtable"			VARCHAR (255), 
	"translatedcolumn"			VARCHAR (255), 
	"rowsource"			VARCHAR (255)
);

-- CREATE INDEXES ...

DROP TABLE IF EXISTS "public"."querydefinitions";
CREATE TABLE IF NOT EXISTS "public"."querydefinitions"
 (
	"queryname"			VARCHAR (255), 
	"sqltext"			TEXT
);

-- CREATE INDEXES ...

DROP TABLE IF EXISTS "public"."judgements";
CREATE TABLE IF NOT EXISTS "public"."judgements"
 (
	"judgement_id"			SERIAL, 
	"sanction"			VARCHAR (255)
);

-- CREATE INDEXES ...
ALTER TABLE "public"."judgements" ADD CONSTRAINT "judgements_pkey" PRIMARY KEY ("judgement_id");


-- CREATE Relationships ...
-- ALTER TABLE "public"."MSysNavPaneGroups" ADD CONSTRAINT "msysnavpanegroups_groupcategoryid_fk" FOREIGN KEY ("groupcategoryid") REFERENCES "public"."MSysNavPaneGroupCategories"("id") ON UPDATE CASCADE ON DELETE CASCADE DEFERRABLE INITIALLY IMMEDIATE;
-- ALTER TABLE "public"."MSysNavPaneGroupToObjects" ADD CONSTRAINT "msysnavpanegrouptoobjects_groupid_fk" FOREIGN KEY ("groupid") REFERENCES "public"."MSysNavPaneGroups"("id") ON UPDATE CASCADE ON DELETE CASCADE DEFERRABLE INITIALLY IMMEDIATE;
-- Relationship from "public"."entries" ("judgement_id") to "public"."judgements"("judgement_id") does not enforce integrity.
-- Relationship from "public"."entries" ("source_id") to "public"."sources"("source_id") does not enforce integrity.
-- Relationship from "public"."entries" ("community_id") to "public"."communities"("community_id") does not enforce integrity.
-- Relationship from "public"."entries" ("land_use_id") to "public"."land_use"("land_use_id") does not enforce integrity.
-- Relationship from "public"."entries" ("actor_id") to "public"."persons"("person_id") does not enforce integrity.
-- Relationship from "public"."entries" ("legal_source_id") to "public"."legal_sources"("legal_source_id") does not enforce integrity.
-- Relationship from "public"."communities" ("parish_id") to "public"."parishes"("parish_id") does not enforce integrity.
-- Relationship from "public"."entries" ("winner_id") to "public"."winners"("winner_id") does not enforce integrity.
-- Relationship from "public"."entries" ("season_id") to "public"."seasons"("season_id") does not enforce integrity.
