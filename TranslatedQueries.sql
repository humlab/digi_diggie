SELECT
    [persons].[full_name],
    [communities].[community_name],
    [entries].[year],
    [entries].[description],
    [land_use].[type],
    [seasons].[season_name],
    [entries].[land_rights_status],
    [entries].[original_placename],
    [community_names].[Ortnamn],
    [sources].[source_abbreviation],
    [entries].[reference_number],
    [community_names].[LOPNR],
    [community_names].[N],
    [community_names].[E],
    [entries].[ID]
FROM ( ( ( ( ( sources
INNER JOIN entries ON [sources].[source_id] = [entries].[source_id] )
INNER JOIN seasons ON [seasons].[season_id] = [entries].[season_id] )
INNER JOIN land_use ON [land_use].[land_use_id] = [entries].[resource_id] )
INNER JOIN communities ON [communities].[community_id] = [entries].[community_id] )
INNER JOIN persons ON [persons].[person_id] = [entries].[actor_id] )
INNER JOIN community_names ON community_names.[ID] = entries.[placename_id];



SELECT [communities].[community_name]
FROM [communities];
