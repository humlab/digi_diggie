
\copy "placenames" ("fid", "geom", "ortnamn", "kvartsruta", "nkoordinat", "ekoordinat", "lanskod", "kommunkod", "detaljtyp", "sprak", "lopnummer", "sockenstadkod", "sockenstadnamn")
from 'resources/ortnamn.csv' (format csv, delimiter E'\t', null '', header true, quote E'\\b');