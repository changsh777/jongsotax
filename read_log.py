import os, sys
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
log = os.path.expanduser(r"~\종소세2026\hometax_result_scraper.log")
with open(log, encoding='utf-8', errors='replace') as f:
    lines = f.readlines()
print(f"총 {len(lines)}줄, 최근 50줄:")
for l in lines[-50:]:
    print(l, end='')
