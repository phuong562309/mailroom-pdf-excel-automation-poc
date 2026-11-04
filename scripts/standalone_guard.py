"""Standalone full-store preservation guard; no application or repository dependencies."""


def assert_no_filtered_view_overwrite(total_source_records_before, total_source_records_after,
                                      *, write_buffer_kind='source_store', operation='upsert',
                                      source_record_ids_before=None, source_record_ids_after=None):
    """Reject projections, shrinking stores and missing/duplicate complete-source IDs."""
    prefix = '[FILTERED_VIEW_OVERWRITE_VETO] '
    if write_buffer_kind != 'source_store':
        raise ValueError(prefix + 'Read-Only Projection cannot write/save/certify a source store')
    if operation not in ('append', 'update', 'upsert', 'write', 'save', 'certify'):
        raise ValueError(prefix + 'Only non-destructive store operations are admitted')
    if any(type(n) is not int or n < 0 for n in (total_source_records_before, total_source_records_after)):
        raise ValueError(prefix + 'Full-source record counts must be nonnegative integers')
    if total_source_records_after < total_source_records_before:
        raise ValueError(prefix + 'Source record count decreased')
    if source_record_ids_before is not None or source_record_ids_after is not None:
        before, after = source_record_ids_before, source_record_ids_after
        if (not isinstance(before, (list, tuple)) or not isinstance(after, (list, tuple))
                or len(before) != total_source_records_before or len(after) != total_source_records_after
                or any(not isinstance(x, str) or not x for x in (*before, *after))
                or len(set(before)) != len(before) or len(set(after)) != len(after)
                or not set(before).issubset(after)):
            raise ValueError(prefix + 'Source record identities lost, duplicated or incomplete')
    return {'status': 'accepted', 'total_source_records_before': total_source_records_before,
            'total_source_records_after': total_source_records_after}
