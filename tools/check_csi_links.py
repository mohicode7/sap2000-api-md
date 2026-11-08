from pathlib import Path
import re,json
from urllib.parse import urlsplit,unquote
for repo in Path('research/converted').iterdir():
 report=json.loads((repo/'validation.json').read_text()); errors=[]
 for p in repo.rglob('*.md'):
  text=p.read_text(encoding='utf-8'); text=re.sub(r'(?ms)^`{3,}[^\n]*\n.*?^`{3,}\s*$', '', text)
  for dest in re.findall(r'\]\(([^\s)]+)(?:\s+"[^"]*")?\)',text):
   u=urlsplit(dest)
   if u.scheme or u.netloc or not u.path: continue
   target=p.parent/unquote(u.path)
   if not target.exists():errors.append({'page':str(p.relative_to(repo)),'target':dest})
 report['missing_local_targets']=errors
 (repo/'validation.json').write_text(json.dumps(report,indent=2))
 print(repo.name,len(errors),errors[:2])
