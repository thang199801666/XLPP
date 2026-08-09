#!/usr/bin/env python3
import argparse, hashlib, json, subprocess, tempfile, zipfile
from pathlib import Path

def fingerprint(path: Path):
    with zipfile.ZipFile(path) as z:
        return {n: hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist()}

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('manifest', type=Path)
    ap.add_argument('--exercise', required=True, type=Path)
    ap.add_argument('--report', type=Path)
    args=ap.parse_args()
    manifest=json.loads(args.manifest.read_text())
    results=[]; root=args.manifest.parent
    for item in manifest['workbooks']:
        src=(root/item['path']).resolve()
        before=fingerprint(src)
        for action, expected in item['expected'].items():
            if not expected: continue
            with tempfile.TemporaryDirectory() as td:
                out=Path(td)/(src.stem+'-'+action+src.suffix)
                proc=subprocess.run([str(args.exercise),str(src),str(out),action],capture_output=True,text=True)
                changed=[]; added=[]; removed=[]
                if proc.returncode==0:
                    after=fingerprint(out)
                    changed=sorted(k for k in before.keys() & after.keys() if before[k]!=after[k])
                    added=sorted(after.keys()-before.keys()); removed=sorted(before.keys()-after.keys())
                allowed_removed=set(item.get('allow_removed_parts', []))
                unexpected_removed=sorted(set(removed)-allowed_removed)
                passed=proc.returncode==0 and not unexpected_removed
                results.append({'file':str(src),'class':item['class'],'action':action,'pass':passed,
                                'changed_parts':changed,'added_parts':added,'removed_parts':removed,
                                'unexpected_removed_parts':unexpected_removed,
                                'stdout':proc.stdout.strip(),'stderr':proc.stderr.strip()})
    report={'schema':1,'results':results,'passed':sum(r['pass'] for r in results),'total':len(results)}
    text=json.dumps(report,indent=2)
    if args.report: args.report.write_text(text)
    print(text)
    return 0 if report['passed']==report['total'] else 1
if __name__=='__main__': raise SystemExit(main())
