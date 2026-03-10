# DigiDiggie Schema Documentation

**Schema:** `digidiggie_tng`

DigiDiggie The Next Generation schema - normalized structure for Swedish court records

## Table of Contents

### Entity Tables

 [community](#community)
 [court_case](#court_case)
 [court_case_entry](#court_case_entry)
 [land_rights_status](#land_rights_status)
 [land_use](#land_use)
 [legal_source](#legal_source)
 [outcome_type](#outcome_type)
 [person](#person)
 [person_entry](#person_entry)
 [person_outcome](#person_outcome)
 [person_relationship](#person_relationship)
 [placename](#placename)
 [role](#role)
 [ruling](#ruling)
 [season](#season)
 [source](#source)

### Lookup Tables

 [parish](#parish)
 [relationship_type](#relationship_type)
 [role_type](#role_type)
 [ruling_type](#ruling_type)

---

## Entity Tables

### community

Social or administrative grouping (e.g., parish or community) to which a person belongs at the time of the court case

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `community_id` | `integer` |  | Primary key for the community table |
| `community_name` | `text` |  | The name of the community (village) |
| `parish_id` | `integer` | ✓ | Foreign key linking to the parish this community belongs to |

### court_case

Single court proceeding recorded in a historical source, identified by date, court, and source text. A court case results in exactly one ruling in this model

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `court_case_id` | `integer` |  | Primary key for the court case table |
| `source_id` | `integer` |  | Foreign key to the source document (e.g., court record collection) |
| `reference_number` | `character varying(16)` | ✓ | Reference number within a specific source collection, e.g., K.B. Wiklund's transcripts |
| `district_court_name` | `text` | ✓ | Name of the district court (tingslag) where the case was heard |
| `case_year` | `integer` | ✓ | The year the case was heard at court |
| `source_text` | `text` | ✓ | Aggregated description/text from the source document for this case |

### court_case_entry

Discrete unit of information extracted from a court case, typically describing a specific land-use situation, event, or claim

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `court_case_entry_id` | `integer` |  | Primary key for the court case entry table |
| `court_case_id` | `integer` |  | Foreign key linking to the parent court case |
| `entry_year` | `integer` | ✓ | The year the event occurred, or if unknown, the year the matter was heard at court |
| `curated_text` | `text` | ✓ | Curated description of the event in free text |
| `original_placename` | `text` | ✓ | The place's name as written in the original source document |
| `season_id` | `integer` | ✓ | Foreign key to the season when the disputed resource was used |
| `land_use_id` | `integer` | ✓ | Foreign key to the resource or land use type involved in the dispute |
| `placename_id` | `integer` | ✓ | Foreign key to a standardized placename from the Swedish National Survey (enables GIS connection) |

### land_rights_status

Description of a person's legal or customary status to the land as interpreted from the entry (e.g., owned land, no land rights, uncertain)

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `land_rights_status_id` | `integer` |  | Primary key for the land rights status table |
| `land_rights_status` | `character varying(255)` |  | Land rights status name (e.g., 'Ja', 'Nej', 'Nja') |
| `description` | `text` |  | Description of the land rights status |

### land_use

Categorized description of how land is used or claimed in an entry (e.g., fishing, hunting, herding, reindeer grazing), based on interpretation of the source

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `land_use_id` | `integer` |  | Primary key for the land use table |
| `description` | `text` |  | The type of land use or resource (e.g., 'Fishing rights', 'Reindeer grazing') |

### legal_source

Normative legal text (law code, regulation, precedent) that a ruling cites or applies

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `legal_source_id` | `integer` |  | Primary key for the legal source table |
| `legal_source_name` | `text` |  | The name of the legal source or legal precedent cited |

### outcome_type

Categorization describing the kind of decision outcome (e.g., winner, sanction, injunction with fine, partition of land)

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `outcome_type_id` | `integer` |  | Primary key for the outcome type table |
| `outcome_type_name` | `text` |  | Outcome type name (e.g., 'Vinnare', 'Böter', 'Friad') |
| `description` | `text` | ✓ | Description of the outcome type |

### person

Historical individual identified in the sources, with personal attributes where known (name, patronymic, birth/death year, notes)

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `person_id` | `integer` |  | Primary key for the person table |
| `given_name` | `text` | ✓ | The person's given name(s) (first name) |
| `patronymic` | `text` | ✓ | The person's patronymic name (e.g., 'Andersson', 'Jonsdotter') |
| `surname` | `text` | ✓ | The person's surname, byname, or family name |
| `birth_year` | `integer` | ✓ | The year of birth |
| `death_year` | `integer` | ✓ | The year of death |
| `community_name` | `text` | ✓ | The name of the village where the person primarily resided |
| `note` | `text` | ✓ | Additional notes about the person |
| `full_name` | `text` | ✓ | The person's full name, computed from given name, patronymic, and surname |

### person_entry

Contextualized appearance of a person within a specific court case entry. Captures the person's role, land rights status, and how they are described in the source

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `person_entry_id` | `integer` |  | Primary key for the person entry table |
| `court_case_entry_id` | `integer` |  | Foreign key linking to the court case entry |
| `person_id` | `integer` |  | Foreign key linking to the person involved |
| `community_id` | `integer` | ✓ | Foreign key to the community where the person resided at the time |
| `land_rights_status_id` | `integer` |  | Foreign key indicating if the person had land rights |
| `role_id` | `integer` | ✓ | Foreign key to the person's role in the case (e.g., plaintiff, defendant, witness) |
| `curated_text` | `text` | ✓ | Additional curated text describing the person's involvement |

### person_outcome

Outcome of a ruling as it affects a specific person (e.g., being sanctioned, declared winner). Connects rulings to individuals

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `person_outcome_id` | `integer` |  | Primary key for the person outcome table |
| `ruling_id` | `integer` |  | Foreign key linking to the ruling |
| `person_id` | `integer` |  | Foreign key linking to the person |
| `outcome_type_id` | `integer` |  | Foreign key to the outcome type (e.g., damages, fined, acquitted, winner) |
| `description` | `text` | ✓ | Description of the specific outcome for this person |

### person_relationship

Relationships between persons (family, social connections)

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `person_relationship_id` | `integer` |  | Primary key for the person relationship table |
| `person_1_id` | `integer` |  | Foreign key to the first person in the relationship |
| `person_2_id` | `integer` |  | Foreign key to the second person in the relationship |
| `relationship_type_id` | `integer` |  | Foreign key to the type of relationship (e.g., father, mother, sibling, spouse) |
| `description` | `text` | ✓ | Additional description of the relationship |

### placename

Standardized geographical place associated with an entry, linked to the Swedish National Survey (external authority/placename registry) with coordinates

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `placename_id` | `integer` |  | Primary key for the placename table |
| `placename` | `text` | ✓ | The standardized placename (ortnamn) |
| `northing` | `integer` | ✓ | Northing coordinate in SWEREF 99 TM (EPSG:3006) |
| `easting` | `integer` | ✓ | Easting coordinate in SWEREF 99 TM (EPSG:3006) |
| `serial_number` | `text` | ✓ | Serial number (löpnummer) from the National Survey |
| `name_type_code` | `text` | ✓ | Name type code (namntyp) from the National Survey |
| `language_code` | `text` | ✓ | Language code (språk) from the National Survey |
| `parish_code` | `text` | ✓ | Parish/town code (sockenstad) from the National Survey |
| `county_code` | `text` | ✓ | County code (län) from the National Survey |
| `municipality_code` | `text` | ✓ | Municipality code (kommun) from the National Survey |
| `combined_placename` | `text` | ✓ | Combined placename (combination of multiple name components) |
| `parish_name` | `text` | ✓ | Parish name from the National Survey |
| `geom` | `geometry(Point,4326)` | ✓ | PostGIS geometry point in WGS84 (EPSG:4326) for mapping |

### role

Social or legal role attributed to a person in a specific entry (e.g., Nybyggare, Sámi, plaintiff, defendant). Roles are contextual, not permanent identities

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `role_id` | `integer` |  | Primary key for the role table |
| `role_name` | `text` |  | Role name (e.g., 'Klagande', 'Svarande', 'Vittne', 'Same', 'Nybyggare') |
| `role_type_id` | `integer` |  | Foreign key to role type (social or judicial) |
| `description` | `text` |  | Description of the role |

### ruling

Judicial decision resulting from a court case. Records the year, description, ruling type (resolved or referred), and may cite a legal source

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `ruling_id` | `integer` |  | Primary key for the ruling table |
| `court_case_id` | `integer` |  | Foreign key linking to the court case (one-to-one relationship) |
| `ruling_year` | `integer` | ✓ | The year the ruling was issued |
| `description` | `text` | ✓ | Description of the ruling in free text |
| `ruling_type_id` | `integer` |  | Foreign key to the type of ruling (e.g., judgement, settlement, referral) |
| `legal_source_id` | `integer` | ✓ | Foreign key to the legal source or precedent cited in the ruling |

### season

Optional temporal qualifier indicating when the event described in an entry took place (e.g., summer, winter)

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `season_id` | `integer` |  | Primary key for the season table |
| `season_name` | `text` |  | The name of the season when the disputed resource was primarily used (e.g., 'Winter', 'Summer') |

### source

Historical source from which court cases are excerpted (e.g., court records, archival volumes), with identifiers and metadata

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `source_id` | `integer` |  | Primary key for the source table |
| `source_name` | `text` |  | The full name of the historical source document |
| `source_abbreviation` | `character varying(255)` | ✓ | The abbreviation for the source, used for quick reference (e.g., 'DB' for court record) |

## Lookup Tables

### parish

Lookup table for parishes

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `parish_id` | `integer` |  | Primary key for the parish table |
| `parish` | `text` |  | Parish name - heading comes from the National Survey database of place names |

### relationship_type

Lookup table for person relationship types

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `relationship_type_id` | `integer` |  | Primary key for the relationship type table |
| `relationship_type_name` | `text` |  | Relationship type name (e.g., 'Far', 'Mor', 'Syskon', 'Make/Maka') |
| `description` | `text` |  | Description of the relationship type |

### role_type

Lookup table for role type categories (social or judicial)

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `role_type_id` | `integer` |  | Primary key for the role type table |
| `role_type_name` | `text` |  | Role type category name ('Social' or 'Juridisk') |
| `description` | `text` |  | Description of the role type category |

### ruling_type

Lookup table for ruling/judgement types

| Column | Type | Nullable | Description |
|--------|------|----------|-------------|
| `ruling_type_id` | `integer` |  | Primary key for the ruling type table |
| `ruling_type` | `character varying(255)` |  | Ruling type name (e.g., 'Dom', 'Förlikning', 'Hänvisning') |
| `description` | `text` |  | Description of the ruling type |
