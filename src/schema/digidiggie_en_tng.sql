
drop schema if exists digidiggie cascade;
create schema digidiggie;
set search_path  to digidiggie, public;

CREATE TABLE "communities" (
  "community_id" serial,
  "community_name" text COLLATE "pg_catalog"."default" NOT NULL DEFAULT ''::text,
  "parish_id" int4,
  CONSTRAINT "communities_pkey" PRIMARY KEY ("community_id")
);
ALTER TABLE "communities" OWNER TO "gudrun";
CREATE INDEX "communities_parish_id_idx" ON "communities" USING btree (
  "parish_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);

CREATE TABLE "court_cases" (
  "court_case_id" serial,
  "source_id" int4 NOT NULL,
  "reference_number" varchar(16) COLLATE "pg_catalog"."default",
  "district_court_name" text COLLATE "pg_catalog"."default",
  "case_date" date,
  "source_text" text COLLATE "pg_catalog"."default",
  CONSTRAINT "court_cases_pkey" PRIMARY KEY ("court_case_id")
);
ALTER TABLE "court_cases" OWNER TO "gudrun";
CREATE INDEX "court_cases_case_date_idx" ON "court_cases" USING btree (
  "case_date" "pg_catalog"."date_ops" ASC NULLS LAST
);
CREATE INDEX "court_cases_reference_number_idx" ON "court_cases" USING btree (
  "reference_number" COLLATE "pg_catalog"."default" "pg_catalog"."text_ops" ASC NULLS LAST
);
CREATE INDEX "court_cases_source_id_idx" ON "court_cases" USING btree (
  "source_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);
CREATE INDEX "entries_reference_number_idx" ON "court_cases" USING btree (
  "reference_number" COLLATE "pg_catalog"."default" "pg_catalog"."text_ops" ASC NULLS LAST
);

CREATE TABLE "entries" (
  "entry_id" serial,
  "court_case_id" int4 NOT NULL,
  "year" int4,
  "description" text COLLATE "pg_catalog"."default",
  "season_id" int4,
  "land_use_id" int4,
  "original_placename" varchar(50) COLLATE "pg_catalog"."default",
  "placename_id" int4,
  CONSTRAINT "entries_pkey" PRIMARY KEY ("entry_id")
);
ALTER TABLE "entries" OWNER TO "gudrun";
CREATE INDEX "entries_court_case_id_idx" ON "entries" USING btree (
  "court_case_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);
CREATE INDEX "entries_land_use_id_idx" ON "entries" USING btree (
  "land_use_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);
CREATE INDEX "entries_placename_id_idx" ON "entries" USING btree (
  "placename_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);
CREATE INDEX "entries_season_id_idx" ON "entries" USING btree (
  "season_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);

CREATE TABLE "judgements" (
  "judgement_id" serial,
  "sanction" text COLLATE "pg_catalog"."default" NOT NULL DEFAULT ''::text,
  CONSTRAINT "judgements_pkey" PRIMARY KEY ("judgement_id")
);
ALTER TABLE "judgements" OWNER TO "gudrun";

CREATE TABLE "land_right_status" (
  "land_rights_status_id" int4 NOT NULL,
  "land_rights_status" varchar(255) NOT NULL,
  PRIMARY KEY ("land_rights_status_id")
);

CREATE TABLE "land_use" (
  "land_use_id" serial,
  "type" text COLLATE "pg_catalog"."default" NOT NULL DEFAULT ''::text,
  CONSTRAINT "land_use_pkey" PRIMARY KEY ("land_use_id")
);
ALTER TABLE "land_use" OWNER TO "gudrun";

CREATE TABLE "legal_sources" (
  "legal_source_id" serial,
  "legal_source_name" text COLLATE "pg_catalog"."default" NOT NULL DEFAULT ''::text,
  CONSTRAINT "legal_sources_pkey" PRIMARY KEY ("legal_source_id")
);
ALTER TABLE "legal_sources" OWNER TO "gudrun";

CREATE TABLE "outcome_types" (
  "outcome_type_id" serial,
  "outcome_type_name" text COLLATE "pg_catalog"."default" NOT NULL,
  "description" text COLLATE "pg_catalog"."default",
  CONSTRAINT "outcome_types_pkey" PRIMARY KEY ("outcome_type_id"),
  CONSTRAINT "outcome_types_outcome_type_name_key" UNIQUE ("outcome_type_name")
);
ALTER TABLE "outcome_types" OWNER TO "gudrun";

CREATE TABLE "parishes" (
  "parish_id" serial,
  "parish" text COLLATE "pg_catalog"."default" NOT NULL DEFAULT ''::text,
  CONSTRAINT "parishes_pkey" PRIMARY KEY ("parish_id")
);
ALTER TABLE "parishes" OWNER TO "gudrun";
CREATE INDEX "parishes_parish_idx" ON "parishes" USING btree (
  "parish" COLLATE "pg_catalog"."default" "pg_catalog"."text_ops" ASC NULLS LAST
);

CREATE TABLE "person_entries" (
  "person_entry_id" serial,
  "entry_id" int4 NOT NULL,
  "actor_id" int4 NOT NULL,
  "community_id" int4,
  "land_rights_status_id" int4 NOT NULL,
  "role_id" int4,
  "curated_text" text COLLATE "pg_catalog"."default",
  CONSTRAINT "person_entries_pkey" PRIMARY KEY ("person_entry_id")
);
ALTER TABLE "person_entries" OWNER TO "gudrun";
CREATE INDEX "person_entries_actor_id_idx" ON "person_entries" USING btree (
  "actor_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);
CREATE INDEX "person_entries_community_id_idx" ON "person_entries" USING btree (
  "community_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);
CREATE INDEX "person_entries_entry_id_idx" ON "person_entries" USING btree (
  "entry_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);

CREATE TABLE "person_outcomes" (
  "person_outcome_id" serial,
  "ruling_id" int4 NOT NULL,
  "person_id" int4 NOT NULL,
  "outcome_type_id" int4 NOT NULL,
  "description" text COLLATE "pg_catalog"."default",
  CONSTRAINT "person_outcomes_pkey" PRIMARY KEY ("person_outcome_id")
);
ALTER TABLE "person_outcomes" OWNER TO "gudrun";
CREATE INDEX "person_outcomes_outcome_type_id_idx" ON "person_outcomes" USING btree (
  "outcome_type_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);
CREATE INDEX "person_outcomes_person_id_idx" ON "person_outcomes" USING btree (
  "person_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);
CREATE INDEX "person_outcomes_ruling_id_idx" ON "person_outcomes" USING btree (
  "ruling_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);

CREATE TABLE "person_properties" (
  "person_property_id" serial,
  "person_id" int4,
  "property_id" int4,
  "specifier" text COLLATE "pg_catalog"."default",
  "property_value" text COLLATE "pg_catalog"."default" NOT NULL DEFAULT ''::text,
  CONSTRAINT "person_properties_pkey" PRIMARY KEY ("person_property_id")
);
ALTER TABLE "person_properties" OWNER TO "gudrun";

CREATE TABLE "persons" (
  "person_id" serial,
  "given_name" text COLLATE "pg_catalog"."default",
  "patronymic" text COLLATE "pg_catalog"."default",
  "surname" text COLLATE "pg_catalog"."default",
  "birth_year" int4,
  "death_year" int4,
  "community_name" text COLLATE "pg_catalog"."default",
  "full_name" text COLLATE "pg_catalog"."default" GENERATED ALWAYS AS (
((((COALESCE(given_name, ''::text) || ' '::text) || COALESCE(patronymic, ''::text)) || ' '::text) || COALESCE(surname, ''::text))
) STORED,
  CONSTRAINT "persons_pkey" PRIMARY KEY ("person_id")
);
ALTER TABLE "persons" OWNER TO "gudrun";

CREATE TABLE "placenames" (
  "fid" int4 NOT NULL,
  "geom" text COLLATE "pg_catalog"."default",
  "ortnamn" text COLLATE "pg_catalog"."default",
  "kvartsruta" text COLLATE "pg_catalog"."default",
  "nkoordinat" float8,
  "ekoordinat" float8,
  "lanskod" text COLLATE "pg_catalog"."default",
  "kommunkod" text COLLATE "pg_catalog"."default",
  "detaljtyp" text COLLATE "pg_catalog"."default",
  "sprak" text COLLATE "pg_catalog"."default",
  "lopnummer" text COLLATE "pg_catalog"."default",
  "sockenstadkod" text COLLATE "pg_catalog"."default",
  "sockenstadnamn" text COLLATE "pg_catalog"."default",
  "geom_point" geometry,
  CONSTRAINT "placenames_pkey" PRIMARY KEY ("fid")
);
ALTER TABLE "placenames" OWNER TO "gudrun";

CREATE TABLE "properties" (
  "property_id" serial,
  "property_name" text COLLATE "pg_catalog"."default",
  "description" text COLLATE "pg_catalog"."default",
  CONSTRAINT "properties_pkey" PRIMARY KEY ("property_id")
);
ALTER TABLE "properties" OWNER TO "gudrun";

CREATE TABLE "roles" (
  "role_id" serial,
  "role_name" text COLLATE "pg_catalog"."default" NOT NULL,
  "description" text COLLATE "pg_catalog"."default",
  CONSTRAINT "roles_pkey" PRIMARY KEY ("role_id"),
  CONSTRAINT "roles_role_name_key" UNIQUE ("role_name")
);
ALTER TABLE "roles" OWNER TO "gudrun";

CREATE TABLE "ruling_type" (
  "ruling_type_id" int4 NOT NULL,
  "ruling_type" varchar(255) NOT NULL,
  PRIMARY KEY ("ruling_type_id")
);

CREATE TABLE "rulings" (
  "ruling_id" serial,
  "court_case_id" int4 NOT NULL,
  "year" int4,
  "description" text COLLATE "pg_catalog"."default",
  "ruling_type_id" int4 NOT NULL,
  "legal_source_id" int4,
  CONSTRAINT "rulings_pkey" PRIMARY KEY ("ruling_id"),
  CONSTRAINT "rulings_court_case_id_key" UNIQUE ("court_case_id")
);
ALTER TABLE "rulings" OWNER TO "gudrun";
CREATE INDEX "rulings_legal_source_id_idx" ON "rulings" USING btree (
  "legal_source_id" "pg_catalog"."int4_ops" ASC NULLS LAST
);

CREATE TABLE "seasons" (
  "season_id" serial,
  "season_name" text COLLATE "pg_catalog"."default" NOT NULL DEFAULT ''::text,
  CONSTRAINT "seasons_pkey" PRIMARY KEY ("season_id")
);
ALTER TABLE "seasons" OWNER TO "gudrun";
CREATE INDEX "seasons_season_name_idx" ON "seasons" USING btree (
  "season_name" COLLATE "pg_catalog"."default" "pg_catalog"."text_ops" ASC NULLS LAST
);

CREATE TABLE "sources" (
  "source_id" serial,
  "source_name" text COLLATE "pg_catalog"."default" NOT NULL DEFAULT ''::text,
  "source_abbreviation" varchar(255) COLLATE "pg_catalog"."default",
  CONSTRAINT "sources_pkey" PRIMARY KEY ("source_id")
);
ALTER TABLE "sources" OWNER TO "gudrun";

CREATE TABLE "winners" (
  "winner_id" serial,
  "winner_description" text COLLATE "pg_catalog"."default" NOT NULL DEFAULT ''::text,
  CONSTRAINT "winners_pkey" PRIMARY KEY ("winner_id")
);
ALTER TABLE "winners" OWNER TO "gudrun";

ALTER TABLE "communities" ADD CONSTRAINT "communities_parish_id_fkey" FOREIGN KEY ("parish_id") REFERENCES "parishes" ("parish_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "court_cases" ADD CONSTRAINT "court_cases_source_id_fkey" FOREIGN KEY ("source_id") REFERENCES "sources" ("source_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "entries" ADD CONSTRAINT "entries_court_case_id_fkey" FOREIGN KEY ("court_case_id") REFERENCES "court_cases" ("court_case_id") ON DELETE CASCADE ON UPDATE NO ACTION;
ALTER TABLE "entries" ADD CONSTRAINT "entries_land_use_id_fkey" FOREIGN KEY ("land_use_id") REFERENCES "land_use" ("land_use_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "entries" ADD CONSTRAINT "entries_placename_id_fkey" FOREIGN KEY ("placename_id") REFERENCES "placenames" ("fid") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "entries" ADD CONSTRAINT "entries_season_id_fkey" FOREIGN KEY ("season_id") REFERENCES "seasons" ("season_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "person_entries" ADD CONSTRAINT "person_entries_actor_id_fkey" FOREIGN KEY ("actor_id") REFERENCES "persons" ("person_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "person_entries" ADD CONSTRAINT "person_entries_community_id_fkey" FOREIGN KEY ("community_id") REFERENCES "communities" ("community_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "person_entries" ADD CONSTRAINT "person_entries_entry_id_fkey" FOREIGN KEY ("entry_id") REFERENCES "entries" ("entry_id") ON DELETE CASCADE ON UPDATE NO ACTION;
ALTER TABLE "person_entries" ADD CONSTRAINT "person_entries_role_id_fkey" FOREIGN KEY ("role_id") REFERENCES "roles" ("role_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "person_entries" ADD CONSTRAINT "fk_person_entries_land_right_status_1" FOREIGN KEY ("land_rights_status_id") REFERENCES "land_right_status" ("land_rights_status_id");
ALTER TABLE "person_outcomes" ADD CONSTRAINT "person_outcomes_outcome_type_id_fkey" FOREIGN KEY ("outcome_type_id") REFERENCES "outcome_types" ("outcome_type_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "person_outcomes" ADD CONSTRAINT "person_outcomes_person_id_fkey" FOREIGN KEY ("person_id") REFERENCES "persons" ("person_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "person_outcomes" ADD CONSTRAINT "person_outcomes_ruling_id_fkey" FOREIGN KEY ("ruling_id") REFERENCES "rulings" ("ruling_id") ON DELETE CASCADE ON UPDATE NO ACTION;
ALTER TABLE "person_properties" ADD CONSTRAINT "person_properties_person_id_fkey" FOREIGN KEY ("person_id") REFERENCES "persons" ("person_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "person_properties" ADD CONSTRAINT "person_properties_property_id_fkey" FOREIGN KEY ("property_id") REFERENCES "properties" ("property_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "rulings" ADD CONSTRAINT "rulings_court_case_id_fkey" FOREIGN KEY ("court_case_id") REFERENCES "court_cases" ("court_case_id") ON DELETE CASCADE ON UPDATE NO ACTION;
ALTER TABLE "rulings" ADD CONSTRAINT "rulings_legal_source_id_fkey" FOREIGN KEY ("legal_source_id") REFERENCES "legal_sources" ("legal_source_id") ON DELETE NO ACTION ON UPDATE NO ACTION;
ALTER TABLE "rulings" ADD CONSTRAINT "fk_rulings_ruling_type_1" FOREIGN KEY ("ruling_type_id") REFERENCES "ruling_type" ("ruling_type_id");

