"""Author the two-sheet mailroom workbook with formula-preserving numeric OOXML caches."""
from collections import Counter
import datetime as dt
import os
from pathlib import Path
import tempfile
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.properties import CalcProperties

HEADERS = ['Doc_Index', 'Document_Type', 'Person_Name', 'Status', 'Document_ID', 'Date',
           'Page_Range', 'Page_Count', 'Output_Filename', 'SharePoint_Relative_Path']
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def put_text(cell, value):
    """Write literal source text without interpreting leading '=' as a formula."""
    cell.value = value
    if isinstance(value, str):
        cell.data_type = 's'


def cache_formulas(filename, caches):
    """Add numeric initial values to existing formula cells; preserve all other ZIP entries."""
    filename = Path(filename)
    ET.register_namespace('', NS)
    fd, temp = tempfile.mkstemp(prefix='.cached-', suffix='.xlsx', dir=filename.parent)
    os.close(fd)
    try:
        with ZipFile(filename, 'r') as original, ZipFile(temp, 'w') as patched:
            for entry in original.infolist():
                content = original.read(entry.filename)
                if entry.filename in caches:
                    root = ET.fromstring(content)
                    remaining = dict(caches[entry.filename])
                    for cell in root.iter(f'{{{NS}}}c'):
                        ref = cell.attrib.get('r')
                        if ref in remaining:
                            if cell.find(f'{{{NS}}}f') is None:
                                raise ValueError(f'Cache target {ref} is not a formula')
                            value = remaining.pop(ref)
                            if type(value) not in (int, float):
                                raise ValueError('Formula caches must be numeric')
                            cell.set('t', 'n')
                            node = cell.find(f'{{{NS}}}v')
                            if node is None:
                                node = ET.SubElement(cell, f'{{{NS}}}v')
                            node.text = str(value)
                    if remaining:
                        raise ValueError(f'Missing formula cache targets: {sorted(remaining)}')
                    content = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                patched.writestr(entry, content)
        os.replace(temp, filename)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)


def write_workbook(manifest, output):
    """Build all summaries from the full admitted records; do not claim to evaluate Excel."""
    records, types, statuses = manifest['records'], manifest['document_types'], manifest['statuses']
    if not records:
        raise ValueError('Cannot create an empty record workbook')
    if len({x.casefold() for x in statuses}) != len(statuses):
        raise ValueError('Status categories differ only by case')
    wb = Workbook()
    summary = wb.active
    summary.title = '11_Type_Summary'
    detail = wb.create_sheet('Extracted_Records')
    wb.calculation = CalcProperties(calcId=191029, calcMode='auto', fullCalcOnLoad=True, forceFullCalc=True)
    wb.properties.creator = 'Mailroom automation proof of concept'
    wb.properties.title = 'Mailroom document breakdown'
    wb.properties.description = 'Synthetic text-layer proof of concept; local outbox only'
    n = len(records)
    last, sum_last = n + 5, len(types) + 5
    total_row, end = sum_last + 1, len(statuses) + 3
    end_letter = get_column_letter(end)
    for j, header in enumerate(HEADERS, 1):
        detail.cell(5, j, header)
    for row, record in enumerate(records, 6):
        for j, header in enumerate(HEADERS, 1):
            value = dt.date.fromisoformat(record[header]) if header == 'Date' else record[header]
            put_text(detail.cell(row, j), value)
        detail.cell(row, 6).number_format = 'mm/dd/yyyy'
        for j in (1, 8):
            detail.cell(row, j).number_format = '#,##0'
    sum_headers = ['Document_Type', 'Document_Count', *statuses, 'Total_Pages']
    for j, header in enumerate(sum_headers, 1):
        put_text(summary.cell(5, j), header)
    counts = Counter(r['Document_Type'] for r in records)
    status_counts = Counter((r['Document_Type'], r['Status']) for r in records)
    page_counts = Counter()
    for record in records:
        page_counts[record['Document_Type']] += record['Page_Count']
    type_range = f"'Extracted_Records'!$B$6:$B${last}"
    status_range = f"'Extracted_Records'!$D$6:$D${last}"
    page_range = f"'Extracted_Records'!$H$6:$H${last}"
    cache = {}
    for row, type_name in enumerate(types, 6):
        summary.cell(row, 1, type_name)
        summary.cell(row, 2, f'=COUNTIFS({type_range},$A{row})')
        cache[f'B{row}'] = counts[type_name]
        for j, status in enumerate(statuses, 3):
            letter = get_column_letter(j)
            criteria = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({letter}$5,"~","~~"),"*","~*"),"?","~?")'
            summary.cell(row, j, f'=COUNTIFS({type_range},$A{row},{status_range},{criteria})')
            cache[f'{letter}{row}'] = status_counts[type_name, status]
        summary.cell(row, end, f'=SUMIFS({page_range},{type_range},$A{row})')
        cache[f'{end_letter}{row}'] = page_counts[type_name]
    summary.cell(total_row, 1, 'TOTAL')
    for j in range(2, end + 1):
        letter = get_column_letter(j)
        summary.cell(total_row, j, f'=SUM({letter}6:{letter}{sum_last})')
        cache[f'{letter}{total_row}'] = sum(cache[f'{letter}{row}'] for row in range(6, sum_last + 1))
    if cache[f'B{total_row}'] != n or cache[f'{end_letter}{total_row}'] != manifest['source_page_count']:
        raise ValueError('Complete-record cache totals disagree with admitted source')
    for j, status in enumerate(statuses, 3):
        if cache[f'{get_column_letter(j)}{total_row}'] != sum(r['Status'] == status for r in records):
            raise ValueError('Complete-record status cache total mismatch')
    summary['A2'] = 'Mailroom document breakdown'
    summary['A3'] = 'Synthetic text-layer sample | Local outbox | No OCR or cloud sync'
    detail['A2'] = 'Extracted mailroom records'
    detail['A3'] = f"Source: {manifest['source']} | SHA-256: {manifest['source_sha256']}"
    dark, muted, line = '183153', '5C6B7B', '8292A6'
    for sheet, right, bottom in ((summary, end, total_row), (detail, 10, last)):
        sheet.sheet_view.showGridLines = False
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.sheet_properties.tabColor = dark if sheet is summary else line
        for cells in sheet.iter_rows(min_row=2, max_row=bottom, min_col=1, max_col=right):
            for cell in cells:
                cell.font = Font(name='Arial', size=11, color='17283C')
                cell.alignment = Alignment(vertical='center', horizontal='left')
        sheet['A2'].font = Font(name='Arial', size=16, bold=True, color=dark)
        sheet['A3'].font = Font(name='Arial', size=11, italic=True, color=muted)
        sheet.row_dimensions[2].height = 28
        sheet.row_dimensions[3].height = 22
        sheet.row_dimensions[5].height = 34
        for row in range(6, bottom + 1):
            sheet.row_dimensions[row].height = 26
        for j in range(1, right + 1):
            sheet.cell(4, j).border = Border(bottom=Side(style='thin', color=line))
            cell = sheet.cell(5, j)
            cell.fill = PatternFill('solid', fgColor=dark)
            cell.font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        sheet.print_options.horizontalCentered = True
        sheet.print_title_rows = '1:5'
        sheet.sheet_properties.outlinePr.summaryRight = False
    for j in range(1, end + 1):
        summary.column_dimensions[get_column_letter(j)].width = 23 if j == 1 else max(19, len(sum_headers[j-1]) + 3)
    for row in range(6, total_row + 1):
        for j in range(2, end + 1):
            summary.cell(row, j).number_format = '#,##0'
            summary.cell(row, j).alignment = Alignment(horizontal='right', vertical='center')
    for j in range(1, end + 1):
        cell = summary.cell(total_row, j)
        cell.fill = PatternFill('solid', fgColor='E6ECF3')
        cell.font = Font(name='Arial', size=11, bold=True, color='17283C')
        cell.border = Border(top=Side(style='thin', color=line))
    summary.row_dimensions[total_row].height = 28
    widths = [11, 18, 26, 22, 22, 17, 15, 15, 65, 103]
    for j, width in enumerate(widths, 1):
        detail.column_dimensions[get_column_letter(j)].width = width
    for row in range(6, last + 1):
        for j in (1, 8):
            detail.cell(row, j).alignment = Alignment(horizontal='right', vertical='center')
    detail.freeze_panes = 'D6'
    for sheet, ref, name in ((summary, f'A5:{end_letter}{sum_last}', 'DocumentTypeSummary'),
                             (detail, f'A5:J{last}', 'ExtractedRecords')):
        table = Table(displayName=name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showFirstColumn=False,
                                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        sheet.add_table(table)
        sheet.auto_filter.ref = ref
    output = Path(output)
    wb.save(output)
    cache_formulas(output, {'xl/worksheets/sheet1.xml': cache})
    # Reopen both formula and cached-value modes; no formula evaluator is claimed.
    formulas = load_workbook(output, read_only=True, data_only=False)
    values = load_workbook(output, read_only=True, data_only=True)
    try:
        for coord, expected in cache.items():
            if not formulas['11_Type_Summary'][coord].data_type == 'f':
                raise ValueError(f'Formula was lost during cache injection: {coord}')
            if values['11_Type_Summary'][coord].value != expected:
                raise ValueError(f'Cached result was lost during export: {coord}')
    finally:
        formulas.close()
        values.close()
    return {'formula_caches': len(cache), 'formula_cache_policy': 'full-record rebuild',
            'formula_evaluator_used': False, 'native_spreadsheet_application_tested': False}
