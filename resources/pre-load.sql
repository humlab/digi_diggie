
alter table placenames alter column geom type text using geom::text;

\copy placenames(fid, geom, ortnamn, kvartsruta, nkoordinat, ekoordinat, lanskod, kommunkod, detaljtyp, sprak, lopnummer, sockenstadkod, sockenstadnamn) FROM 'resources/ortnamn.csv' WITH (FORMAT csv, HEADER true);

alter table placenames
  add column geom_point geometry(point, 3006);  -- or 4326, depending on crs

update placenames
  set geom_point = st_setsrid(st_makepoint(ekoordinat, nkoordinat), 3006);


/* example use:
SELECT ST_AsGeoJSON(ST_Transform(geom, 4326)) FROM placenames;
*/