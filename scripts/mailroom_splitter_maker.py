"""Split contracted text-layer mailroom PDFs, validate fully, then publish one local batch."""
import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile
import unicodedata
from urllib.parse import unquote
from pypdf import PdfReader, PdfWriter

BASE = Path(__file__).resolve().parents[1]
try:
    from guardrail_checks import assert_no_filtered_view_overwrite
    GUARD_BACKEND = 'shared'
except ImportError:
    if __package__:
        from .standalone_guard import assert_no_filtered_view_overwrite
    else:
        from standalone_guard import assert_no_filtered_view_overwrite
    GUARD_BACKEND = 'standalone'
if __package__:
    from .workbook_writer import write_workbook
else:
    from workbook_writer import write_workbook


def portable_ref(path):
    """Store repository-relative paths, or only external-source basename plus its hash."""
    path = Path(path).resolve()
    try:
        return path.relative_to(BASE).as_posix()
    except ValueError:
        return path.name


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def filename_component(value):
    """Reversible on NFC strings, including separator underscores and literal percent signs."""
    value = unicodedata.normalize('NFC', value)
    encoded = []
    tail = len(value.rstrip(' .'))
    for i, char in enumerate(value):
        if char in '%_<>:"/\\|?*' or ord(char) < 32 or ord(char) == 127 or (i >= tail and char in ' .'):
            encoded.append(''.join(f'%{b:02X}' for b in char.encode('utf-8')))
        else:
            encoded.append(char)
    result = ''.join(encoded)
    if not result or unquote(result, errors='strict') != value:
        raise ValueError('Cannot encode filename component reversibly')
    return result


def parse(reader, types):
    records, seen_ids = [], set()
    active = None
    field_names = {'Document Type': 'Document_Type', 'Document ID': 'Document_ID',
                   'Person Name': 'Person_Name', 'Status': 'Status', 'Date': 'Date'}
    for idx, page in enumerate(reader.pages, 1):
        text = page.extract_text() or ''
        if not text.strip():
            raise ValueError(f'Page {idx}: image-only or empty; OCR not supported')
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        fields = {}
        for label, key in field_names.items():
            values = [line[len(label) + 1:].strip() for line in lines if line.startswith(label + ':')]
            if len(values) > 1 or (values and not values[0]):
                raise ValueError(f'Page {idx}: ambiguous/empty {label}')
            if values:
                if unicodedata.normalize('NFC', values[0]) != values[0]:
                    raise ValueError(f'Page {idx}: {label} must be NFC Unicode')
                fields[key] = values[0]
        markers = [re.fullmatch(r'Page ([1-9]\d*) of ([1-9]\d*)', line) for line in lines]
        markers = [m for m in markers if m]
        if len(markers) != 1:
            raise ValueError(f'Page {idx}: require exactly one document-local Page x of y')
        current, total = map(int, markers[0].groups())
        if current > total or not all(k in fields for k in ('Document_Type', 'Document_ID')):
            raise ValueError(f'Page {idx}: invalid sequence or missing type/ID')
        if fields['Document_Type'] not in types:
            raise ValueError(f'Page {idx}: unknown type {fields["Document_Type"]}')
        if current == 1:
            if active is not None:
                raise ValueError(f'Page {idx}: new document before preceding sequence complete')
            missing = set(field_names.values()) - set(fields)
            if missing:
                raise ValueError(f'Page {idx}: missing first-page fields {sorted(missing)}')
            if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', fields['Date']):
                raise ValueError(f'Page {idx}: Date must be YYYY-MM-DD')
            dt.date.fromisoformat(fields['Date'])
            if fields['Document_ID'].casefold() in seen_ids:
                raise ValueError(f'Page {idx}: duplicate Document ID')
            seen_ids.add(fields['Document_ID'].casefold())
            active = dict(fields, Doc_Index=len(records) + 1, start_page=idx, Page_Count=total, source_pages=[])
        elif active is None:
            raise ValueError(f'Page {idx}: orphan continuation')
        if current != len(active['source_pages']) + 1 or total != active['Page_Count']:
            raise ValueError(f'Page {idx}: incomplete/conflicting page sequence')
        for key, value in fields.items():
            if value != active[key]:
                raise ValueError(f'Page {idx}: conflicting {key}')
        active['source_pages'].append(idx)
        if current == total:
            active['Page_Range'] = str(active['start_page']) if total == 1 else f'{active["start_page"]}-{idx}'
            records.append(active)
            active = None
    if active is not None or not records:
        raise ValueError('Incomplete final document or no records')
    if [p for r in records for p in r['source_pages']] != list(range(1, len(reader.pages) + 1)):
        raise ValueError('Source page coverage is incomplete')
    return records


def run(source, output, config):
    source, output = Path(source).resolve(), Path(output).resolve()
    if output.exists():
        raise FileExistsError(f'Refusing to overwrite existing output directory: {portable_ref(output)}')
    cfg = json.loads(Path(config).read_text(encoding='utf-8'))
    types = cfg['document_types']
    if len(types) != 11 or len(set(types)) != 11 or any(not re.fullmatch(r'Type_\d{2}', t) for t in types):
        raise ValueError('Configuration must supply eleven unique safe Type_XX labels')
    source_hash = digest(source)
    reader = PdfReader(source)
    if reader.is_encrypted:
        raise ValueError('Encrypted PDFs are unsupported')
    records = parse(reader, types)
    statuses = sorted({r['Status'] for r in records})
    if len({s.casefold() for s in statuses}) != len(statuses):
        raise ValueError('Ambiguous status categories differ only by case; use one consistent spelling per status')
    used_paths = set()
    for r in records:
        r['Output_Filename'] = '_'.join(filename_component(r[k]) for k in ['Person_Name', 'Status', 'Document_ID']) + '.pdf'
        if len(r['Output_Filename'].encode('utf-16-le')) // 2 > cfg['max_filename_utf16_units']:
            raise ValueError(f'{r["Document_ID"]}: encoded filename too long for Windows')
        r['SharePoint_Relative_Path'] = f'SharePoint_Synced_Outbox/{r["Document_Type"]}/{r["Output_Filename"]}'
        if len(str(output / r['SharePoint_Relative_Path']).encode('utf-16-le')) // 2 > 32700:
            raise ValueError('Windows extended path budget exceeded; choose a shallower output directory')
        key = r['SharePoint_Relative_Path'].casefold()
        if key in used_paths:
            raise ValueError(f'Case-insensitive output collision: {r["Output_Filename"]}')
        used_paths.add(key)
    output.parent.mkdir(parents=True, exist_ok=True)
    stage = Path(tempfile.mkdtemp(prefix=f'.{output.name}.stage-', dir=output.parent))
    try:
        for r in records:
            target = stage / r['SharePoint_Relative_Path']
            target.parent.mkdir(parents=True, exist_ok=True)
            writer = PdfWriter()
            for p in r['source_pages']:
                writer.add_page(reader.pages[p - 1])
            with target.open('xb') as f:
                writer.write(f)
            r['pdf_sha256'] = digest(target)
        manifest = {'schema_version': cfg['schema_version'], 'source': portable_ref(source), 'source_sha256': source_hash,
                    'document_types': types, 'statuses': statuses,
                    'document_count': len(records), 'source_page_count': len(reader.pages), 'records': records,
                    'filename_policy': cfg['filename_policy'], 'local_outbox_only': True, 'ocr_performed': False}
        manifest_path = stage / 'maker_manifest.json'
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        workbook_info = write_workbook(manifest, stage / '11_Types_Summary_Breakdown.xlsx')
        if digest(source) != source_hash:
            raise ValueError('Source changed during processing')
        before_pages = [f'page:{p}' for p in range(1, len(reader.pages) + 1)]
        after_pages = [f'page:{p}' for r in records for p in r['source_pages']]
        page_guard = assert_no_filtered_view_overwrite(len(before_pages), len(after_pages), operation='write',
                         source_record_ids_before=before_pages, source_record_ids_after=after_pages)
        record_ids = [r['Document_ID'] for r in records]
        record_guard = assert_no_filtered_view_overwrite(len(record_ids), len(records), operation='write',
                           source_record_ids_before=record_ids,
                           source_record_ids_after=[r['Document_ID'] for r in manifest['records']])
        build = {'maker_status': 'complete_unverified', 'documents': len(records), 'pages': len(reader.pages),
                 'source_sha256': source_hash, 'workbook_sha256': digest(stage / '11_Types_Summary_Breakdown.xlsx'),
                 'manifest_sha256': digest(manifest_path), 'independent_audit_required': True,
                 'guard_backend': GUARD_BACKEND, 'workbook_export': workbook_info,
                 'page_write_guard': page_guard, 'record_write_guard': record_guard,
                 'max_output_path_utf16_units': max(len(str(output / r['SharePoint_Relative_Path']).encode('utf-16-le'))//2 for r in records)}
        (stage / 'maker_build.json').write_text(json.dumps(build, indent=2) + '\n', encoding='utf-8')
        # Serialize commits from cooperating CLI processes on Windows and POSIX.
        lock = output.parent / f'.{output.name}.publish.lock'
        with lock.open('xb'):
            pass
        try:
            if output.exists():
                raise FileExistsError(f'Output appeared during build: {portable_ref(output)}')
            os.rename(stage, output)
        finally:
            lock.unlink(missing_ok=True)
        print(json.dumps(dict(build, output=portable_ref(output))))
        return build
    except BaseException:
        # Only the private stage we just created may be removed.
        if stage.exists() and stage.parent == output.parent and stage.name.startswith(f'.{output.name}.stage-'):
            shutil.rmtree(stage)
        raise


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--input', required=True, type=Path, help='Combined PDF conforming to mailroom-text-v1')
    ap.add_argument('--output-dir', '--output', dest='output', required=True, type=Path, help='New batch directory (must not exist)')
    ap.add_argument('--config', type=Path, default=BASE / 'config.json')
    args = ap.parse_args()
    try:
        run(args.input, args.output, args.config)
    except Exception as exc:
        # Do not leak host home/repository paths into portable saved run logs.
        message = str(exc).replace(str(BASE), '[repository]').replace(str(Path.home()), '[home]')
        print(f'MAILROOM_REJECTED: {message}', file=sys.stderr)
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(main())
