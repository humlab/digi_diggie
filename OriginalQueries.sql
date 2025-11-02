/*SELECT
    [Personer].[Helnamn],
    [Byar].[Bynamn],
    [Entries].[År],
    [Entries].[Beskrivning],
    [Markanvändning].[Typ],
    [Årstid].[Årstid],
    [Entries].[OK],
    [Entries].[Ortnamn_orig],
    [Ortnamn_ny].[Ortnamn],
    [Källor].[Kortkälla],
    [Entries].[Refnr],
    [Ortnamn_ny].[LOPNR],
    [Ortnamn_ny].[N],
    [Ortnamn_ny].[E],
    [Entries].[ID]
FROM
    (
        (
            (
                (
                    (
                        Källor
                        INNER JOIN Entries ON [Källor].[ID] = [Entries].[Källa]
                    )
                    INNER JOIN Årstid ON [Årstid].[ID] = [Entries].[Årstid2]
                )
                INNER JOIN Markanvändning ON [Markanvändning].[ID] = [Entries].[Resurs]
            )
            INNER JOIN Byar ON [Byar].[ID] = [Entries].[By]
        )
        INNER JOIN Personer ON [Personer].[ID] = [Entries].[Aktör]
    )
    INNER JOIN Ortnamn_ny ON Ortnamn_ny.[ID] = Entries.[Ortnamn];
*/

SELECT
    [persons].[Helnamn],
    [Byar].[Bynamn],
    [Entries].[År],
    [Entries].[Beskrivning],
    [Markanvändning].[Typ],
    [Årstid].[Årstid],
    [Entries].[OK],
    [Entries].[Ortnamn_orig],
    [Ortnamn_ny].[Ortnamn],
    [Källor].[Kortkälla],
    [Entries].[Refnr],
    [Ortnamn_ny].[LOPNR],
    [Ortnamn_ny].[N],
    [Ortnamn_ny].[E],
    [Entries].[ID]
FROM
    (
        (
            (
                (
                    (
                        Källor
                        INNER JOIN Entries ON [Källor].[ID] = [Entries].[Källa]
                    )
                    INNER JOIN Årstid ON [Årstid].[ID] = [Entries].[Årstid2]
                )
                INNER JOIN Markanvändning ON [Markanvändning].[ID] = [Entries].[Resurs]
            )
            INNER JOIN Byar ON [Byar].[ID] = [Entries].[By]
        )
        INNER JOIN Personer ON [Personer].[ID] = [Entries].[Aktör]
    )
    INNER JOIN Ortnamn_ny ON Ortnamn_ny.[ID] = Entries.[Ortnamn];



SELECT
    Byar.Bynamn
FROM
    Byar;

    -- [Källor]
    -- INNER JOIN (
    --     Årstid
    --     INNER JOIN (
    --         Markanvändning
    --         INNER JOIN (
    --             Byar
    --             INNER JOIN (
    --                 Personer
    --                 INNER JOIN (
    --                     Ortnamn_ny
    --                     INNER JOIN Entries ON Ortnamn_ny.[ID] = Entries.[Ortnamn]
    --                 ) ON Personer.ID = Entries.Aktör
    --             ) ON Byar.ID = Entries.By
    --         ) ON Markanvändning.ID = Entries.Resurs
    --     ) ON Årstid.ID = Entries.Årstid2
    -- ) ON Källor.ID = Entries.Källa;
    