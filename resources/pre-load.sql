-- run drop_all first to ensure clean state

---------
drop table if exists public.placenames cascade;
-- create table if not exists "placenames"
--  (
--     "fid" integer primary key,
--     "geom" text,
--     "ortnamn" text,
--     "kvartsruta" text,
--     "nkoordinat" double precision,
--     "ekoordinat" double precision,
--     "lanskod" text,
--     "kommunkod" text,
--     "detaljtyp" text,
--     "sprak" text,
--     "lopnummer" text,
--     "sockenstadkod" text,
--     "sockenstadnamn" text
-- );
-- ---------

-- create table if not exists "placenames"
--  (
--     "ID" serial primary key,
--     "Ortnamn" text,
--     "N" double precision,
--     "E" double precision,
--     "LOPNR" text,
--     "Namntyp_nr" text,
--     "Språk_nr" text,
--     "Sockenstad_nr" text,
--     "Län_nr" text,
--     "Kommun_nr" text,
--     "Kombo" text,
--     "Sockenstad" text,
--     "nr" text    
-- );


create table if not exists "placenames"
 (
    "id" serial primary key,
    "ortnamn" text,
    "n" double precision,
    "e" double precision,
    "lopnr" text,
    "namntyp_nr" text,
    "språk_nr" text,
    "sockenstad_nr" text,
    "lan_nr" text,
    "kommun_nr" text,
    "kombo" text,
    "sockenstad" text,
    "nr" text    
);

--alter table placenames alter column geom type text using geom::text;

--\copy placenames(fid, geom, ortnamn, kvartsruta, nkoordinat, ekoordinat, lanskod, kommunkod, detaljtyp, sprak, lopnummer, sockenstadkod, sockenstadnamn) FROM 'resources/placenames.csv' WITH (FORMAT csv, HEADER true, ENCODING 'WIN1252');
--\copy placenames("Ortnamn", "N", "E", "LOPNR", "Namntyp_nr", "Språk_nr", "Sockenstad_nr", "Län_nr", "Kommun_nr", "Kombo", "Sockenstad", "nr") FROM 'resources/placenames.csv' WITH (FORMAT csv, HEADER true, ENCODING 'WIN1252');
\copy placenames(ortnamn, n, e, lopnr, namntyp_nr, språk_nr, sockenstad_nr, lan_nr, kommun_nr, kombo, sockenstad, nr) FROM 'resources/placenames.csv' WITH (FORMAT csv, HEADER true, ENCODING 'WIN1252');

--alter table placenames
  add column geom_point geometry(point, 3006);  -- or 4326, depending on crs

--update placenames
  set geom_point = st_setsrid(st_makepoint(ekoordinat, nkoordinat), 3006);


/* example use:
SELECT ST_AsGeoJSON(ST_Transform(geom, 4326)) FROM placenames;
*/