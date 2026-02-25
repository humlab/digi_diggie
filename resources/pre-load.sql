-- run drop_all first to ensure clean state

drop table if exists placenames cascade;

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

\copy placenames(id, ortnamn, n, e, lopnr, namntyp_nr, språk_nr, sockenstad_nr, lan_nr, kommun_nr, kombo, sockenstad, nr) FROM 'resources/placenames.csv' WITH (FORMAT csv, HEADER true, ENCODING 'WIN1252');