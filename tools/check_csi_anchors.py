from pathlib import Path
import json,re,posixpath
from urllib.parse import urlsplit,unquote
for repo in Path('research/converted').iterdir():
 m=json.loads((repo/'manifest.json').read_text()); anchors={t['markdown'].lower():set(t['anchors']) for t in m['topics']}; errors=[]
 for topic in m['topics']:
  p=repo/topic['markdown']; text=p.read_text(encoding='utf-8'); text=re.sub(r'(?ms)^`{3,}[^\n]*\n.*?^`{3,}\s*$','',text)
  for url in re.findall(r'\]\(([^\s)]+)\)',text):
   u=urlsplit(url)
   if u.scheme or not u.fragment:continue
   dest=posixpath.normpath(posixpath.join(posixpath.dirname(topic['markdown']),unquote(u.path))) if u.path else topic['markdown']
   if dest.lower() in anchors and unquote(u.fragment) not in anchors[dest.lower()]:errors.append((topic['markdown'],url))
 print(repo.name,'fragment issues',len(errors),errors[:3])
