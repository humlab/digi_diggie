    -- insert into digidiggie_tng.person_entries (
    --     entry_id, 
    --     actor_id, 
    --     community_id, 
    --     land_rights_status_id,
    --     role_id,
    --     curated_text
    -- )
    select 
        e.entry_id,
        coalesce(e.actor_id, 0) as actor_id, -- actor_id refers to person_id
        e.community_id,
        coalesce(lrs.land_rights_status_id, 
                (select land_rights_status_id from digidiggie_tng.land_right_status where land_rights_status = 'unknown' limit 1)
        ) as land_rights_status_id,
        null as role_id, -- FIXME: #18 Add role_id mapping if possible, otherwise will need manual population. No direct mapping in old schema.
        null as curated_text -- FIXME: #19 Add curated_text if possible, otherwise will need manual population. No direct mapping in old schema.
    from digidiggie_tog.entries e
    left join digidiggie_tng.land_right_status lrs on 
        lrs.land_rights_status = coalesce(e.land_rights_status, 'unknown')
    where e.entry_id in (select entry_id from digidiggie_tng.entries)
    and e.actor_id is not null;


