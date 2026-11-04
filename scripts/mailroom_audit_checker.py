#!/usr/bin/env python3
"""Independent, source-derived mailroom audit. No Maker/oracle imports or reads."""
import argparse
import collections
import datetime as dt
import hashlib
import io
import json
import pathlib
import re
import os
import sys
import unicodedata

import openpyxl
import pdfplumber
from openpyxl.utils.cell import range_boundaries, get_column_letter
try:
    from guardrail_checks import assert_no_filtered_view_overwrite
    GUARD_BACKEND='external'
except ImportError:
    from standalone_guard import assert_no_filtered_view_overwrite
    GUARD_BACKEND='standalone'

TYPES = tuple(f'Type_{i:02}' for i in range(1, 12))
HEADERS = ['Doc_Index','Document_Type','Person_Name','Status','Document_ID','Date',
           'Page_Range','Page_Count','Output_Filename','SharePoint_Relative_Path']

def sha(path):
    return hashlib.sha256(pathlib.Path(path).read_bytes()).hexdigest()

def relative(path,base):
    return pathlib.PurePath(os.path.relpath(path,base)).as_posix()

def require(ok, message):
    if not ok:
        raise ValueError(message)

def component(value):
    """Independent implementation of the published reversible component grammar."""
    value = unicodedata.normalize('NFC', value)
    result = []
    trailing = len(value.rstrip(' .'))
    for i, char in enumerate(value):
        if char in '%_<>:"/\\|?*' or ord(char) < 32 or ord(char) == 127 or (i >= trailing and char in ' .'):
            result.extend(f'%{byte:02X}' for byte in char.encode('utf-8'))
        else:
            result.append(char)
    return ''.join(result)

def uncomponent(value):
    raw = bytearray()
    i = 0
    while i < len(value):
        if value[i] == '%':
            require(bool(re.fullmatch('[0-9A-F]{2}', value[i+1:i+3])), 'Invalid percent escape')
            raw.append(int(value[i+1:i+3], 16)); i += 3
        else:
            raw.extend(value[i].encode('utf-8')); i += 1
    return raw.decode('utf-8')

def filename(record):
    return '_'.join(component(record[k]) for k in ('Person_Name','Status','Document_ID')) + '.pdf'

def table_evidence(page):
    """Recognize bordered tables and tables with horizontal rules/aligned text.

    Metadata/prose is not enough: candidates need at least three rows with
    multiple occupied cells and repeated vertical text alignment across rows.
    No sample coordinates, header names or expected cell values are used.
    """
    grid=page.find_tables()
    aligned=page.find_tables({'vertical_strategy':'text','horizontal_strategy':'lines','min_words_vertical':3})
    def substantial(table):
        rows=table.extract()
        return len(rows)>=3 and all(sum(bool((cell or '').strip()) for cell in row)>=2 for row in rows)
    return {'grid_tables':sum(substantial(t) for t in grid),
            'horizontal_text_tables':sum(substantial(t) for t in aligned)}

def read_pdf(path):
    pages = []
    with pdfplumber.open(path) as pdf:
        require(bool(pdf.pages), 'Empty PDF')
        for number, page in enumerate(pdf.pages, 1):
            text = (page.extract_text(x_tolerance=2, y_tolerance=3) or '').strip()
            require(bool(text), f'Image-only/empty page {number}; OCR unsupported')
            labels = {}
            for label, key in [('Document Type','Document_Type'),('Person Name','Person_Name'),
                               ('Status','Status'),('Document ID','Document_ID'),('Date','Date')]:
                found = re.findall(r'^' + re.escape(label) + r':\s*(.*?)\s*$', text, re.M)
                require(len(found) <= 1, f'Ambiguous {label} on page {number}')
                if found:
                    require(bool(found[0]), f'Empty {label} on page {number}')
                    labels[key] = found[0]
            local = re.findall(r'^Page\s+(\d+)\s+of\s+(\d+)\s*$', text, re.M)
            require(len(local) == 1, f'Missing/ambiguous page sequence on page {number}')
            x, y = map(int, local[0])
            require(1 <= x <= y, f'Invalid sequence on page {number}')
            require(labels.get('Document_Type') in TYPES and labels.get('Document_ID'),
                    f'Missing/unknown identity on page {number}')
            tables=table_evidence(page)
            pages.append({'number':number,'local':x,'total':y,'labels':labels,'text':text,
                          'text_sha256':hashlib.sha256(text.encode()).hexdigest(),
                          'geometry':(float(page.width),float(page.height)),
                          'line_count':len(page.lines),'table_count':max(tables.values()),'table_evidence':tables})
    return pages

def source_records(pages):
    records = []
    ids = set()
    cursor = 0
    while cursor < len(pages):
        first = pages[cursor]
        require(first['local'] == 1, f'Orphan continuation at source page {cursor+1}')
        labels = first['labels']
        require(set(labels) == {'Document_Type','Person_Name','Status','Document_ID','Date'},
                f'Incomplete first-page metadata at page {cursor+1}')
        require(re.fullmatch(r'\d{4}-\d{2}-\d{2}',labels['Date']), 'Date must be ISO YYYY-MM-DD')
        dt.date.fromisoformat(labels['Date'])
        require(labels['Document_ID'] not in ids, 'Repeated document identifier')
        ids.add(labels['Document_ID'])
        count = first['total']
        require(cursor+count <= len(pages), 'Truncated final document')
        for offset in range(count):
            page = pages[cursor+offset]
            require((page['local'],page['total']) == (offset+1,count), 'Incomplete/ambiguous page boundary')
            for key,value in page['labels'].items():
                require(value == labels[key], f'Conflicting continuation {key} on page {page["number"]}')
        record = dict(labels, Doc_Index=len(records)+1, Page_Count=count,
                      Page_Range=str(cursor+1) if count == 1 else f'{cursor+1}-{cursor+count}',
                      source_page_numbers=list(range(cursor+1,cursor+count+1)))
        record['Output_Filename'] = filename(record)
        require(len(record['Output_Filename'].encode('utf-16-le'))//2 <= 255, 'Windows filename too long')
        for key,part in zip(('Person_Name','Status','Document_ID'),record['Output_Filename'][:-4].split('_')):
            require(uncomponent(part) == unicodedata.normalize('NFC',record[key]), 'Filename not reversible')
        record['SharePoint_Relative_Path'] = 'SharePoint_Synced_Outbox/'+record['Document_Type']+'/'+record['Output_Filename']
        records.append(record)
        cursor += count
    require(len({r['SharePoint_Relative_Path'].casefold() for r in records}) == len(records), 'Casefold filename collision')
    return records

def raster_hashes(path):
    result=[]
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            rgb=page.to_image(resolution=40,antialias=False).original.convert('RGB')
            result.append(hashlib.sha256(str(rgb.size).encode()+rgb.tobytes()).hexdigest())
    return result

def arguments(text):
    result=[]; start=0; quoted=False; depth=0
    for i,c in enumerate(text):
        if c == '"': quoted=not quoted
        elif not quoted and c == '(':depth+=1
        elif not quoted and c == ')':depth-=1
        elif c == ',' and not quoted and depth==0:
            result.append(text[start:i].strip()); start=i+1
    result.append(text[start:].strip())
    return result

def criterion_matches(value,criterion):
    """Excel equality criteria: * and ? wildcards, ~ escapes the next character."""
    criterion=str(criterion);pattern='';i=0
    while i<len(criterion):
        char=criterion[i]
        if char=='~' and i+1<len(criterion):
            i+=1;pattern+=re.escape(criterion[i])
        elif char=='*':pattern+='.*'
        elif char=='?':pattern+='.'
        else:pattern+=re.escape(char)
        i+=1
    return bool(re.fullmatch(pattern,str(value),re.I|re.S))

def evaluate(workbook, sheet, expression, stack=(), overrides=None):
    if not hasattr(workbook,'_mailroom_formula_cells'):
        workbook._mailroom_formula_cells={(ws.title,cell.coordinate):(cell.value,cell.data_type)
            for ws in workbook for row in ws.iter_rows() for cell in row if cell.value is not None}
    if isinstance(expression,(int,float)):
        return expression
    require(isinstance(expression,str),'Unsupported formula argument')
    expression=expression.strip()
    if expression.startswith('='): expression=expression[1:]
    if expression.startswith('"') and expression.endswith('"'): return expression[1:-1].replace('""','"')
    if re.fullmatch(r'-?\d+(\.\d+)?',expression): return float(expression)
    function=re.fullmatch(r'(COUNTIFS|SUMIFS|SUM|SUBSTITUTE)\((.*)\)',expression,re.I)
    if function:
        name=function[1].upper(); args=arguments(function[2])
        values=[evaluate(workbook,sheet,a,stack,overrides) for a in args]
        if name == 'SUBSTITUTE':
            require(len(values)==3 and all(isinstance(v,str) for v in values),'Unsupported SUBSTITUTE arguments')
            return values[0].replace(values[1],values[2])
        if name == 'SUM':
            return sum(v for val in values for v in (val if isinstance(val,list) else [val]) if isinstance(v,(int,float)))
        sums=values.pop(0) if name == 'SUMIFS' else None
        require(len(values)%2 == 0 and len(values)>=2,'Bad COUNTIFS/SUMIFS arity')
        ranges=values[::2]; criteria=values[1::2]
        require(all(isinstance(r,list) and len(r)==len(ranges[0]) for r in ranges),'Mismatched criteria ranges')
        return sum((sums[i] if sums is not None else 1) for i in range(len(ranges[0]))
                   if all(criterion_matches(r[i],c) for r,c in zip(ranges,criteria)))
    target=sheet
    if '!' in expression:
        target,expression=expression.rsplit('!',1); target=target.strip("'")
    expression=expression.replace('$','')
    require(bool(re.fullmatch(r'[A-Z]+\d+(?::[A-Z]+\d+)?',expression,re.I)), 'Unsupported formula grammar: '+expression)
    if ':' in expression:
        a,b,c,d=range_boundaries(expression)
        values=[]
        for row in range(b,d+1):
            for col in range(a,c+1):
                key=(target,f'{get_column_letter(col)}{row}')
                original,kind=workbook._mailroom_formula_cells.get(key,(None,'n'))
                value=(overrides or {}).get(key,original)
                values.append(evaluate(workbook,target,value,stack,overrides) if kind=='f' else value)
        return values
    key=(target,expression)
    require(key not in stack,'Formula cycle')
    original,kind=workbook._mailroom_formula_cells.get(key,(None,'n'))
    value=(overrides or {}).get(key,original)
    return evaluate(workbook,target,value,stack+(key,),overrides) if kind == 'f' else value

def check_workbook(path, records):
    # In-memory read-only ZIP handles also prevent failed probes retaining a
    # Windows filesystem lock while an exception traceback is still alive.
    data=pathlib.Path(path).read_bytes()
    formulas=openpyxl.load_workbook(io.BytesIO(data),read_only=True,data_only=False)
    cached=openpyxl.load_workbook(io.BytesIO(data),read_only=True,data_only=True)
    try:
        return _check_workbook_open(formulas,cached,records)
    finally:
        formulas.close();cached.close()

def _check_workbook_open(formulas,cached,records):
    require(formulas.sheetnames == ['11_Type_Summary','Extracted_Records'],'Unexpected workbook sheets/order')
    details=formulas['Extracted_Records']; summary=formulas['11_Type_Summary']; cache=cached['11_Type_Summary']
    require([details.cell(5,col).value for col in range(1,11)] == HEADERS,'Detail headers mismatch')
    for row,record in enumerate(records,6):
        for col,key in enumerate(HEADERS,1):
            actual=details.cell(row,col).value
            if key == 'Date':
                require(isinstance(actual,(dt.date,dt.datetime)),f'Date is not typed at F{row}')
                actual=actual.date().isoformat() if isinstance(actual,dt.datetime) else actual.isoformat()
            require(actual == record[key],f'Excel record mismatch {key} row{row}: {actual!r} != {record[key]!r}')
    for row in details.iter_rows(min_row=6+len(records),max_col=10):
        require(not any(cell.value is not None for cell in row),'Extra detail rows')
    statuses=sorted({r['Status'] for r in records})
    require(len({s.casefold() for s in statuses})==len(statuses),'Case-variant statuses merge in Excel criteria')
    headers=['Document_Type','Document_Count']+statuses+['Total_Pages']
    require([summary.cell(5,c).value for c in range(1,len(headers)+1)] == headers,'Dynamic status headers mismatch')
    formula_checks=[]
    for row,typ in enumerate(TYPES,6):
        require(summary.cell(row,1).value == typ,'Missing/reordered configured type')
        group=[r for r in records if r['Document_Type']==typ]
        expected=[len(group)]+[sum(r['Status']==s for r in group) for s in statuses]+[sum(r['Page_Count'] for r in group)]
        for col,value in enumerate(expected,2):
            cell=summary.cell(row,col)
            family='SUMIFS' if col==len(headers) else 'COUNTIFS'
            require(cell.data_type=='f' and re.match('='+family+r'\(',cell.value,re.I),f'Missing {family} formula {cell.coordinate}')
            computed=evaluate(formulas,summary.title,cell.value)
            require(computed==value and cache.cell(row,col).value==value,f'Formula/cache mismatch {cell.coordinate}')
            formula_checks.append({'cell':cell.coordinate,'formula':cell.value,'computed':computed,'cached':cache.cell(row,col).value})
    totalrow=6+len(TYPES)
    require(str(summary.cell(totalrow,1).value).upper()=='TOTAL','Missing TOTAL row')
    expected=[len(records)]+[sum(r['Status']==s for r in records) for s in statuses]+[sum(r['Page_Count'] for r in records)]
    for col,value in enumerate(expected,2):
        cell=summary.cell(totalrow,col)
        require(cell.data_type=='f' and re.match(r'=SUM\(',cell.value,re.I),f'Missing SUM total {cell.coordinate}')
        require(evaluate(formulas,summary.title,cell.value)==value and cache.cell(totalrow,col).value==value,'TOTAL formula/cache mismatch')
        formula_checks.append({'cell':cell.coordinate,'formula':cell.value,'computed':value,'cached':cache.cell(totalrow,col).value})
    for ws in cached:
        for row_index,row in enumerate(ws.iter_rows(),1):
            for col_index,cell in enumerate(row,1):
                require(cell.data_type != 'e',f'Excel error {ws.title}!{get_column_letter(col_index)}{row_index}')
    # Non-mutating dependency probe: alter each source record's type/status/pages
    # in an evaluator overlay. Read-only workbook bytes and caches stay unchanged.
    for index,record in enumerate(records,6):
        newtype=TYPES[(TYPES.index(record['Document_Type'])+1)%len(TYPES)]
        newstatus=statuses[(statuses.index(record['Status'])+1)%len(statuses)]
        overlay={('Extracted_Records',f'B{index}'):newtype,
                 ('Extracted_Records',f'D{index}'):newstatus,
                 ('Extracted_Records',f'H{index}'):record['Page_Count']+1}
        modified=[dict(r) for r in records]
        modified[index-6].update(Document_Type=newtype,Status=newstatus,Page_Count=record['Page_Count']+1)
        for row,typ in enumerate(TYPES,6):
            group=[r for r in modified if r['Document_Type']==typ]
            expected=[len(group)]+[sum(r['Status']==s for r in group) for s in statuses]+[sum(r['Page_Count'] for r in group)]
            for col,value in enumerate(expected,2):
                expression=formulas._mailroom_formula_cells[(summary.title,f'{get_column_letter(col)}{row}')][0]
                require(evaluate(formulas,summary.title,expression,overrides=overlay)==value,
                        f'Formula does not respond to detail row {index} at {summary.cell(row,col).coordinate}')
    return {'statuses':statuses,'formula_checks':formula_checks,'record_rows':len(records),'summary_type_rows':len(TYPES),
            'readonly_formula_dependency_probes':len(records)}

def audit(source,output):
    source=pathlib.Path(source).resolve(); output=pathlib.Path(output).resolve()
    pages=read_pdf(source); records=source_records(pages)
    source_rasters=raster_hashes(source)
    require(len(source_rasters)==len(pages),'Source raster coverage incomplete')
    expected_paths={r['SharePoint_Relative_Path'] for r in records}
    actual_paths={p.relative_to(output).as_posix() for p in (output/'SharePoint_Synced_Outbox').rglob('*.pdf')}
    require(actual_paths==expected_paths,'Unexpected/missing child PDF paths')
    mapped=[]; units=[]; refs=[]
    for record in records:
        child=output/record['SharePoint_Relative_Path']
        require(child.resolve().is_relative_to(output),'Child path escapes outbox')
        childpages=read_pdf(child); own=source_records(childpages)
        require(len(own)==1,'Child contains multiple documents')
        require(len(childpages)==record['Page_Count'],'Child page count mismatch')
        rasters=raster_hashes(child)
        require(len(rasters)==len(childpages),'Child raster coverage incomplete')
        identities=[]
        for i,number in enumerate(record['source_page_numbers']):
            expected=pages[number-1]; actual=childpages[i]
            require(actual['text']==expected['text'] and actual['geometry']==expected['geometry'],f'Page text/geometry mismatch source#{number}')
            require(rasters[i]==source_rasters[number-1],f'Page raster mismatch source#{number}')
            identity=f'{sha(source)}:page:{number}:{expected["text_sha256"]}'
            mapped.append(identity); identities.append(identity)
        refs.append({'path':relative(child,output),'sha256':sha(child)})
        units.append({'document_id':record['Document_ID'],'source_pages':record['source_page_numbers'],
                      'source_page_ids':identities,'child':record['SharePoint_Relative_Path'],'page_count':len(childpages),
                      'metadata_checked':True,'text_checked':True,'raster_checked':True,'filename_checked':True})
    before=[f'{sha(source)}:page:{p["number"]}:{p["text_sha256"]}' for p in pages]
    guard=assert_no_filtered_view_overwrite(len(before),len(mapped),operation='certify',
          source_record_ids_before=before,source_record_ids_after=mapped)
    require(mapped==before,'Source pages not mapped exactly once in source order')
    workbook=output/'11_Types_Summary_Breakdown.xlsx'
    excel=check_workbook(workbook,records)
    refs.append({'path':relative(workbook,output),'sha256':sha(workbook)})
    return {'status':'PASS','source_kind':'labeled text-layer PDF; provenance supplied by caller','source':{'path':relative(source,output),'sha256':sha(source)},
            'documents':len(records),'source_pages':len(pages),'child_pages':len(mapped),'types':len(TYPES),
            'configured_types':len(TYPES),'observed_types':len({r['Document_Type'] for r in records}),
            'checked':len(records)+len(pages),'unchecked':0,'mismatched':0,'failures':0,
            'oracle_access':False,'guard_backend':GUARD_BACKEND,'reference_base':'--output directory',
            'strategy':'pdfplumber/pdfminer direct-source metadata+sequence; PDFium all-page RGB raster identity; openpyxl read-only formula+cache reconciliation',
            'source_page_ids':before,'coverage_guard':guard,'units':units,'excel':excel,
            'source_features':{'document_page_counts':dict(collections.Counter(r['Page_Count'] for r in records)),
             'table_pages':sum(p['table_count']>0 for p in pages),'unicode_names':sum(any(ord(c)>127 for c in r['Person_Name']) for r in records),
             'repeat_people_different_status':any(len({r['Status'] for r in records if r['Person_Name']==n})>1 for n in {r['Person_Name'] for r in records})},
            'output_artifacts':refs,'limitations':['Labeled text-layer input; no OCR or real client-scan accuracy claim.','Local outbox; no remote SharePoint transfer.','Independent PDFium raster comparison at 40 dpi verifies page preservation; it is not a high-resolution typography review.','Supported formulas/caches checked independently; no native spreadsheet-application execution.']}

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source',required=True); parser.add_argument('--output',required=True)
    parser.add_argument('--audit',required=True)
    args=parser.parse_args()
    destination=pathlib.Path(args.audit)
    if destination.exists():
        print(json.dumps({'status':'FAIL','error':'Audit exists; refusing silent overwrite'})); return 2
    try:
        result=audit(args.source,args.output); code=0
    except Exception as exc:
        result={'status':'FAIL','failures':1,'unchecked':None,'mismatched':None,'error':str(exc),'oracle_access':False}; code=1
    destination.parent.mkdir(parents=True,exist_ok=True)
    with destination.open('x',encoding='utf-8') as handle: json.dump(result,handle,ensure_ascii=False,indent=2)
    print(json.dumps({'status':result['status'],'audit':destination.name,'error':result.get('error')},ensure_ascii=False))
    return code

if __name__ == '__main__':
    sys.exit(main())
