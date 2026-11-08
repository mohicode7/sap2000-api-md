from pathlib import Path
import json,re
from urllib.parse import unquote,urlsplit
import sys
sys.path.insert(0,'research/chm-tools/lib')
from bs4 import BeautifulSoup
root=Path('research/converted')
for repo in root.iterdir():
 if not (repo/'manifest.json').exists(): continue
 manifest=json.loads((repo/'manifest.json').read_text())
 errors=[]; code_errors=[]
 for item in manifest['topics']:
  p=repo/item['markdown']; raw=p.read_bytes().replace(b'\r\r\n',b'\n').replace(b'\r\n',b'\n'); p.write_bytes(raw)
  text=raw.decode()
  fences=re.findall(r'^(`{3,})$',text,re.M)
  src=Path('API/_extracted')/(manifest['provenance']['product']+'-'+manifest['provenance']['installed_product_version'])/item['source']
  data=src.read_bytes(); encoding='cp1252' if b'windows-1252' in data[:2000].lower() else 'utf-8-sig'
  doc=BeautifulSoup(data.decode(encoding),'html.parser'); source_codes=doc.select('pre, p.Coding'); item['code_blocks']=len(source_codes)
  if len(fences)%2 or (manifest['provenance']['product']=='etabs' and len(fences)!=len(source_codes)*2): code_errors.append(item['source'])
  for block in source_codes:
   for br in block.find_all('br'): br.replace_with('\n')
   code=block.get_text().replace('\xa0',' ') if 'Coding' in block.get('class',[]) else block.get_text()
   code=code.replace('\r\n','\n').strip('\n')
   if code not in text: code_errors.append(item['source']+': code text differs')
  link_text=re.sub(r'(?ms)^`{3,}[^\n]*\n.*?^`{3,}\s*$', '', text)
  for dest in re.findall(r'\]\(([^\s)]+)(?:\s+"[^"]*")?\)',link_text):
   u=urlsplit(dest)
   if u.scheme or u.netloc or not u.path: continue
   target=p.parent/unquote(u.path)
   if not target.exists(): errors.append({'page':item['markdown'],'target':dest})
 (repo/'manifest.json').write_text(json.dumps(manifest,indent=2))
 report={'topics':len(manifest['topics']),'missing_local_targets':errors,'code_fence_count_mismatches':code_errors,'unresolved_source_links':manifest['provenance']['unresolved_source_links']}
 (repo/'validation.json').write_text(json.dumps(report,indent=2))
 print(repo.name,json.dumps({k:len(v) if isinstance(v,list) else v for k,v in report.items()}))
