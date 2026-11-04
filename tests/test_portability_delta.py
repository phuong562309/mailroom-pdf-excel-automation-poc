#!/usr/bin/env python3
"""One standalone portability Delta; standard-library runner, no test framework."""
import argparse
import ast
import hashlib
import importlib.metadata
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

from reportlab.pdfgen import canvas

REPO=pathlib.Path(__file__).resolve().parents[1]
PACKAGES={'pypdf','pdfplumber','openpyxl','reportlab','pillow'}

def need(value,message):
    if not value:raise ValueError(message)

def sha(path):return hashlib.sha256(pathlib.Path(path).read_bytes()).hexdigest()

def ref(path,root):return {'path':pathlib.Path(path).resolve().relative_to(root.resolve()).as_posix(),'sha256':sha(path)}

def environment():
    env=dict(os.environ)
    for key in list(env):
        if key.upper() in ['PYTHONPATH','PYTHONHOME']:env.pop(key)
    env['PATH']='';env['PYTHONNOUSERSITE']='1';env['PYTHONDONTWRITEBYTECODE']='1'
    env['PYTHONUTF8']='1'
    return env

def run(root,args,success=True,trace=False,cwd=None):
    command=[sys.executable,'-X','utf8','-B']+(['-X','importtime'] if trace else [])+args
    result=subprocess.run(command,cwd=cwd or root,env=environment(),capture_output=True,text=True,encoding='utf-8',timeout=180)
    need((result.returncode==0)==success,'CLI result mismatch: '+str(args[0])+' '+result.stdout[-700:]+' '+result.stderr[-700:])
    need('unrecognized arguments' not in result.stderr and 'the following arguments are required' not in result.stderr,'CLI argument contract mismatch')
    return result

def scan_package(root):
    """Content and import declarations only; no Maker implementation is imported."""
    forbidden=['@'+'oai/'+'artifact-tool','codex'+'-runtimes']
    blocked_dirs={'node'+'_modules','.'+'codex','.'+'git','Foun'+'dry'}
    host=re.compile(r'(?i)(?:[a-z]:[\\/](?:users|documents and settings)[\\/]|/(?:Users|home)/)')
    requirement=root/'requirements.txt'
    names={re.split(r'[<>=!~;\[]',line.strip(),1)[0].lower() for line in requirement.read_text(encoding='utf-8').splitlines() if line.strip() and not line.lstrip().startswith('#')}
    need(PACKAGES.issubset(names),'Five public Python requirements must be declared')
    local={p.stem for p in (root/'scripts').glob('*.py')}
    external={'PIL':'pillow','pypdf':'pypdf','pdfplumber':'pdfplumber','openpyxl':'openpyxl','reportlab':'reportlab'}
    inspected=0;xml_parts=0;import_roots=set()
    for path in root.rglob('*'):
        need(not path.is_symlink() and not (getattr(path.stat(),'st_file_attributes',0)&1024),'Linked filesystem content is not a portable package')
        if path.is_dir():
            need(path.name not in blocked_dirs,'Private/dependency directory included');continue
        payloads=[]
        if path.suffix.lower()=='.xlsx':
            with zipfile.ZipFile(path) as archive:
                payloads=[archive.read(name).decode('utf-8',errors='ignore') for name in archive.namelist() if name.endswith(('.xml','.rels'))]
            xml_parts+=len(payloads)
        elif path.suffix.lower() in ['.py','.md','.txt','.json','.toml','.mjs','.js','.csv','.yaml','.yml','.html']:
            payloads=[path.read_text(encoding='utf-8')]
        else:payloads=[path.read_bytes().decode('utf-8',errors='ignore')]
        for text in payloads:
            need(not host.search(text),'Personal absolute path in '+path.relative_to(root).as_posix())
            need(not any(token.casefold() in text.casefold() for token in forbidden),'Private dependency in '+path.relative_to(root).as_posix())
        if path.suffix=='.py':
            tree=ast.parse(path.read_text(encoding='utf-8'))
            for node in ast.walk(tree):
                mods=[a.name.split('.')[0] for a in node.names] if isinstance(node,ast.Import) else [node.module.split('.')[0]] if isinstance(node,ast.ImportFrom) and node.module else []
                for module in mods:
                    import_roots.add(module)
                    if module in sys.stdlib_module_names or module in local:continue
                    if module=='guardrail_checks':
                        need(any(isinstance(t,ast.Try) and any(isinstance(h.type,ast.Name) and h.type.id=='ImportError' for h in t.handlers)
                                 and any(n is node for body in t.body for n in ast.walk(body)) for t in ast.walk(tree)), 'External guard import lacks ImportError fallback')
                    else:need(module in external and external[module] in names,'Undeclared runtime import: '+module)
            if path.name in ['mailroom_splitter_maker.py','mailroom_audit_checker.py']:
                need(not any(isinstance(n,ast.Call) and isinstance(n.func,ast.Attribute) and n.func.attr=='insert' and
                     isinstance(n.func.value,ast.Attribute) and n.func.value.attr=='path' for n in ast.walk(tree)),'Runtime path injection found')
        inspected+=1
    return {'files':inspected,'xlsx_xml_parts':xml_parts,'requirements':sorted(names),'import_roots':sorted(import_roots),'host_paths':0,'private_dependencies':0,'symlinks':0}

def guard_probe(root):
    code="""
import importlib.util,json,sys
import standalone_guard as g
import mailroom_audit_checker as c
assert importlib.util.find_spec('guardrail_checks') is None
assert c.GUARD_BACKEND=='standalone' and c.assert_no_filtered_view_overwrite.__module__=='standalone_guard'
assert 'guardrail_checks' not in sys.modules
valid=g.assert_no_filtered_view_overwrite(2,3,source_record_ids_before=['a','b'],source_record_ids_after=['a','b','c'])
cases=[(2,1,{}),(2,2,{'source_record_ids_before':['a','b'],'source_record_ids_after':['a','c']}),
(2,2,{'source_record_ids_before':['a','b'],'source_record_ids_after':['a','a']}),
(2,2,{'write_buffer_kind':'projection'}),(2,2,{'operation':'delete'}),(True,2,{})]
for before,after,kw in cases:
 try:g.assert_no_filtered_view_overwrite(before,after,**kw)
 except (ValueError,TypeError):pass
 else:raise AssertionError('Invalid guard input accepted')
print(json.dumps({'standalone_selected':True,'external_guard_absent':True,'negative_cases':len(cases),'valid_growth':True}))
"""
    result=run(root,['-c',code],cwd=root/'scripts')
    return json.loads(result.stdout)

def batch(root,source,output,auditname):
    made=run(root,['scripts/mailroom_splitter_maker.py','--input',source,'--output-dir',output],trace=True)
    need('standalone_guard' in made.stderr,'Maker standalone fallback was not imported in traced CLI')
    run(root,['scripts/mailroom_audit_checker.py','--source',source,'--output',output,'--audit',auditname])
    result=json.loads((root/auditname).read_text(encoding='utf-8'))
    need(result['status']=='PASS' and result['unchecked']==result['mismatched']==result['failures']==0,'Independent batch audit incomplete')
    need(result['guard_backend']=='standalone','Checker did not select packaged fallback')
    for artifact in result['output_artifacts']:
        path=root/output/artifact['path'];need(sha(path)==artifact['sha256'],'Audited output hash drift')
    need(not pathlib.PurePath(result['source']['path']).is_absolute(),'Audit source path is absolute')
    return result

def mini_source(path,statuses):
    c=canvas.Canvas(str(path))
    for index,status in enumerate(statuses,1):
        for line,text in enumerate([f'Document Type: Type_{index:02}',f'Person Name: Example Person {index}',
          f'Status: {status}',f'Document ID: LITERAL-{index}','Date: 2026-11-04','Page 1 of 1']):
            c.drawString(40,740-line*25,text)
        c.showPage()
    c.save()

def adverse(root):
    reports=[]
    source='examples/Combined_Mailroom_Scan_Sample.pdf'
    sentinel=root/'existing-destination';sentinel.mkdir();(sentinel/'KEEP.txt').write_text('keep',encoding='utf-8')
    run(root,['scripts/mailroom_splitter_maker.py','--input',source,'--output','existing-destination'],success=False)
    need((sentinel/'KEEP.txt').read_text()=='keep' and len(list(sentinel.iterdir()))==1,'Existing destination altered')
    reports.append('no silent overwrite')
    blank=root/'blank.pdf';c=canvas.Canvas(str(blank));c.showPage();c.save()
    run(root,['scripts/mailroom_splitter_maker.py','--input','blank.pdf','--output','invalid-output'],success=False)
    need(not (root/'invalid-output').exists(),'Invalid input published a destination')
    reports.append('image-only rejected without publication')
    mini_source(root/'literal-status.pdf',['A*B','A?B','A~B','Normal'])
    literal=batch(root,'literal-status.pdf','literal-output','literal-audit.json')
    need(literal['documents']==4 and literal['observed_types']==4 and literal['configured_types']==11,'Subset batch is not reusable')
    need(literal['excel']['statuses']==sorted(['A*B','A?B','A~B','Normal']),'Literal statuses merged')
    reports.append('literal wildcard statuses and subset N/M formulas/caches')
    mini_source(root/'case-status.pdf',['Ready','ready'])
    run(root,['scripts/mailroom_splitter_maker.py','--input','case-status.pdf','--output','case-output'],success=False)
    need(not (root/'case-output').exists(),'Case-variant status input published')
    reports.append('case-variant status categories rejected')
    return {'checks':reports,'literal_documents':literal['documents'],'literal_formula_cells':len(literal['excel']['formula_checks']),
            'literal_dependency_probes':literal['excel']['readonly_formula_dependency_probes']}

def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--review-metadata');parser.add_argument('--independent-review')
    args=parser.parse_args()
    need(not (REPO/'AUDIT.json').exists() and not (REPO/'PORTABILITY_REPORT.json').exists(),'Acceptance outputs already exist; refusing silent overwrite')
    try:
        initial_scan=scan_package(REPO)
        original=batch(REPO,'examples/Combined_Mailroom_Scan_Sample.pdf','sample-output','AUDIT.json')
        need((original['documents'],original['source_pages'],original['observed_types'])==(15,24,11),'Sample scale mismatch')
        with tempfile.TemporaryDirectory(prefix='mailroom portable é ') as directory:
            relocated=pathlib.Path(directory)/'PoC Unicode Việt space'
            shutil.copytree(REPO,relocated)
            need(relocated!=REPO and ' ' in str(relocated) and any(ord(c)>127 for c in str(relocated)),'Relocation path does not exercise Unicode and spaces')
            standalone=guard_probe(relocated)
            copied=batch(relocated,'examples/Combined_Mailroom_Scan_Sample.pdf','relocated-output','relocated-audit.json')
            need((copied['documents'],copied['source_pages'],copied['observed_types'])==(15,24,11),'Relocated coverage mismatch')
            adverse_cases=adverse(relocated)
            relocated_scan=scan_package(relocated)
        report={'status':'PASS','synthetic':True,'documents':15,'pages':24,'configured_types':11,'observed_types':11,
          'checked':original['checked']+copied['checked'],'unchecked':0,'failures':0,'oracle_access':False,
          'original_directory_run':'PASS','relocated_directory_run':'PASS','relocation_path_classes':['fresh temporary directory','spaces','Unicode'],
          'sanitized_environment':{'path_empty':True,'pythonpath_removed':True,'pythonhome_removed':True,'user_site_disabled':True},
          'standalone_guard':standalone,'adverse_controls':adverse_cases,'initial_scan':initial_scan,'relocated_scan':relocated_scan,
          'formula_cache_cells_per_main_batch':len(original['excel']['formula_checks']),
          'formula_dependency_probes_per_main_batch':original['excel']['readonly_formula_dependency_probes'],
          'renderer':'PDFium through public pdfplumber; no external executables',
          'versions':{p:importlib.metadata.version(p) for p in sorted(PACKAGES)},
          'artifacts':[ref(REPO/'AUDIT.json',REPO),ref(REPO/'sample-output/11_Types_Summary_Breakdown.xlsx',REPO)],
          'limitations':original['limitations']}
        with (REPO/'PORTABILITY_REPORT.json').open('x',encoding='utf-8') as h:json.dump(report,h,ensure_ascii=False,indent=2)
        final_scan=scan_package(REPO)
        if args.review_metadata or args.independent_review:
            need(args.review_metadata and args.independent_review,'Both internal review paths are required')
            metadata=json.loads(pathlib.Path(args.review_metadata).read_text(encoding='utf-8'));project=pathlib.Path(metadata.pop('project_root')).resolve()
            proofref=metadata.get('governance_review_ref')
            need(isinstance(proofref,dict),'Internal acceptance requires a separate governance review reference')
            proofpath=project/proofref['path'];need(sha(proofpath)==proofref['sha256'],'Governance review hash mismatch')
            proof=json.loads(proofpath.read_text(encoding='utf-8'))
            need(proof.get('status')=='PASS' and proof.get('unchecked')==0 and proof.get('failures')==0 and proof.get('checked',0)>=1,
                 'Separate governance review is incomplete')
            need(proof.get('guard_sha256')==metadata.get('governance_guard_sha256'),'Governance review guard hash mismatch')
            units={
             'portable_workbook':('Private authoring engines prevent client execution','Public Python authoring must preserve formulas/caches','Independently compare all rows, statuses and cached formulas'),
             'standalone_runtime':('Ancestor imports and hidden executables break relocation','Ship fallback invariants and use public Python dependencies','Run a fresh relocated copy with sanitized environment'),
             'independent_audit':('Maker replay can conceal record and page loss','Use source-derived pdfminer and independent PDFium checks','Check every page and record with zero unchecked units'),
             'client_package':('Reports and documents can leak host paths','Scan exported payloads and requirement declarations','Package relative references and declared public dependencies'),
             'governance_portability':('Local success alone does not establish a portable deliverable','Apply a reusable portability guard before acceptance','Bind standalone and package-content evidence to the current delivery')}
            review=dict(metadata,status='pass',oracle_access=False,checked=report['checked'],unchecked=0,failures=0,
              strategy='Independent source pdfplumber/pdfminer, PDFium raster, openpyxl formula/cache audit; sanitized original+relocated CLI runs; content/import and fallback guard controls',
              maker_implementation=ref(REPO/'scripts/mailroom_splitter_maker.py',project),checker_implementation=ref(REPO/'scripts/mailroom_audit_checker.py',project),
              batch_artifact=ref(REPO/'examples/Combined_Mailroom_Scan_Sample.pdf',project),
              output_artifacts=[ref(REPO/'AUDIT.json',project),ref(REPO/'PORTABILITY_REPORT.json',project)]+
                [ref(REPO/'sample-output'/r['path'],project) for r in original['output_artifacts']],
              units=[dict(node=k,status='pass',checked=proof['checked'] if k=='governance_portability' else 1,
                     unchecked=0,failures=0,phenomenon=v[0],design_gap=v[1],general_rule=v[2]) for k,v in units.items()])
            with pathlib.Path(args.independent_review).open('x',encoding='utf-8') as h:json.dump(review,h,ensure_ascii=False,indent=2)
        print(json.dumps({'status':'PASS','documents':15,'pages':24,'unchecked':0,'failures':0,'files_scanned':final_scan['files'],
                          'audit':ref(REPO/'AUDIT.json',REPO),'report':ref(REPO/'PORTABILITY_REPORT.json',REPO)},ensure_ascii=False));return 0
    except Exception as exc:
        # Keep transient host locations out of a deliverable failure report.
        message=str(exc).replace(str(REPO),'<repository>')
        failure={'status':'FAIL','error':message,'failures':1}
        print(json.dumps(failure,ensure_ascii=False));return 1

if __name__=='__main__':sys.exit(main())
