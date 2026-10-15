"""Function coverage: Excel's function list (test/excel-functions.txt) vs. functions the app registers."""
import json, sys
impl = set(json.load(open('test/implemented-functions.json')))
tot = miss = 0; rep = []
for line in open('test/excel-functions.txt'):
    if line.startswith('#') or ':' not in line: continue
    cat, names = line.split(':', 1); names = names.split(); m = [n for n in names if n not in impl]
    tot += len(names); miss += len(m)
    rep.append(f"{cat}: {len(names)-len(m)}/{len(names)}" + (f"  MISSING: {' '.join(m)}" if m else ''))
out = f"FUNCTION COVERAGE vs Excel list: {tot-miss}/{tot}\n" + "\n".join(rep)
print(out); open(sys.argv[1], 'w').write(out + '\n')
