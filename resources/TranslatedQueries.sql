select
    [persons].[full_name],
    [communities].[community_name],
    [entries].[year],
    [entries].[description],
    [land_use].[type],
    [seasons].[season_name],
    [entries].[land_rights_status],
    [entries].[original_placename],
    [community_names].[ortnamn],
    [sources].[source_abbreviation],
    [entries].[reference_number],
    [community_names].[lopnr],
    [community_names].[n],
    [community_names].[e],
    [entries].[id]
from ( ( ( ( ( sources
inner join entries on [sources].[source_id] = [entries].[source_id] )
inner join seasons on [seasons].[season_id] = [entries].[season_id] )
inner join land_use on [land_use].[land_use_id] = [entries].[resource_id] )
inner join communities on [communities].[community_id] = [entries].[community_id] )
inner join persons on [persons].[person_id] = [entries].[actor_id] )
inner join community_names on community_names.[id] = entries.[placename_id];

select [communities].[community_name]
from [communities];
